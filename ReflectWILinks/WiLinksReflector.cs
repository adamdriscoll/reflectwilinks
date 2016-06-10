/****************************** Module Header ******************************\
Module Name:  WiLinksReflector
Project:      ReflectWILinks
Copyright (c) Vincent Labatut.

The main engine of the ReflectWILinks utility.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Client;
using System.Diagnostics;
using Microsoft.TeamFoundation.VersionControl.Client;

namespace ReflectWILinks
{
    /// <summary>
    /// Not threadsafe.
    /// </summary>
    public class WiLinksReflector : IDisposable
    {
        private TfsTeamProjectCollection _sourceTpc;
        private TfsTeamProjectCollection _targetTpc;
        private WorkItemStore _sourceWis;
        private WorkItemStore _targetWis;
        private HashSet<int> _processedWorkItems = new HashSet<int>();

        private VersionControlServer _sourceVersionControlServer;
        private VersionControlServer _targetVersionControlServer;


        private const int MAX_WI_LINKS_TO_RESTORE = 30;

        // Key is the ReflectedWorkItemId of the target work items
        // Since target work items are copied from the source work items
        // Key is the Id of the source work item and value the target work item Id
        private Dictionary<int, int> _reflectedWorkItemIds;

        private const string ReflectedIdFieldName = "TfsMigrationToolReflectedWorkItemId";

        public delegate void LogMessage(string message, TraceLevel level);
        public event LogMessage LogMessageEvent;

        public Uri SourceTfsUri { get; private set; }
        public Uri TargetTfsUri { get; private set; }

        public bool AddMissingRelatedLinks { get; set; }
        public bool AddMissingChangesetsLinks { get; set; }
        public bool AddMissingExternalLinks { get; set; }

        public int NbProcessedWorkItems { get; private set; }
        public int NbSourceRelatedLinksFound { get; private set; }
        public int NbSourceExternalLinksFound { get; private set; }
        public int NbSourceChangesetLinksFound { get; private set; }
        public int NbTargetRelatedLinksFound { get; private set; }
        public int NbTargetExternalLinksFound { get; private set; }
        public int NbTargetChangesetLinksFound { get; private set; }
        public int NbTargetRelatedLinksAdded { get; private set; }
        public int NbTargetExternalLinksAdded { get; private set; }
        public int NbTargetChangesetLinksAdded { get; private set; }
        public int NbTargetCrossRelatedLinks { get; private set; }
        public int NbMissingRelatedWorkItems { get; private set; }
        public int NbSaveErrors { get; private set; }
        public string TargetProject { get; private set; }

        public WiLinksReflector(Uri sourceTfsUri, Uri targetTfsUri, string targetProject)
        {
            SourceTfsUri = sourceTfsUri;
            TargetTfsUri = targetTfsUri;
            AddMissingRelatedLinks = true;
            AddMissingChangesetsLinks = true;
            AddMissingExternalLinks = true;
            TargetProject = targetProject;
        }

        /// <summary>
        /// Connects to all Tfs
        /// Forces pre-load of reflected work item ids (can be long for thousands of WI)
        /// </summary>
        public void Initialize()
        {
            log(TraceLevel.Info, "Connecting to {0}", SourceTfsUri);
            _sourceTpc = new TfsTeamProjectCollection(SourceTfsUri);
            _sourceTpc.Authenticate();
            _sourceWis = _sourceTpc.GetService<WorkItemStore>();
            _sourceVersionControlServer = _sourceTpc.GetService<VersionControlServer>();

            log(TraceLevel.Info, "Connecting to {0}", TargetTfsUri);
            _targetTpc = new TfsTeamProjectCollection(TargetTfsUri);
            _targetTpc.Authenticate();

            _targetWis = _targetTpc.GetService<WorkItemStore>();
            _targetVersionControlServer = _targetTpc.GetService<VersionControlServer>();
        }

        private static Guid FindQuery(QueryFolder folder, string queryName)
        {
            foreach (var item in folder)
            {
                if (item.Name.Equals(queryName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return item.Id;
                }

                var itemFolder = item as QueryFolder;
                if (itemFolder != null)
                {
                    var result = FindQuery(itemFolder, queryName);
                    if (!result.Equals(Guid.Empty))
                    {
                        return result;
                    }
                }
            }
            return Guid.Empty;
        }

        public void LoadReflectedWorkItemIds(string reflectedLinksScopeQuery)
        {
            if (_sourceTpc == null) Initialize();

            var id = FindQuery(_targetWis.Projects[TargetProject].QueryHierarchy, reflectedLinksScopeQuery);

            if (id == Guid.Empty)
            {
                log(TraceLevel.Error, "Unable to find query {0}", reflectedLinksScopeQuery);
                return;
            }

            log(TraceLevel.Info, "Found query '{0}' with id '{1}'", reflectedLinksScopeQuery, id);

            LoadReflectedWorkItemIds(id);
        }

        public void LoadReflectedWorkItemIds(Guid reflectedLinksScopeQuery)
        {
            if (_sourceTpc == null) Initialize();

            log(TraceLevel.Info, "Caching all reflected work item Ids...");
            _reflectedWorkItemIds = new Dictionary<int, int>();

            var query = _targetWis.GetQueryDefinition(reflectedLinksScopeQuery);

            log(TraceLevel.Info, "Executing query " + query.Name);
            log(TraceLevel.Verbose, query.QueryText);
            foreach (WorkItem wi in _targetWis.Query(query.QueryText.Replace("@project", "'" + TargetProject + "'")))
            {
                // key is the source work item id
                try
                {
                    if (wi.Fields.Contains(ReflectedIdFieldName))
                    {
                        _reflectedWorkItemIds.Add(int.Parse((string)wi.Fields[ReflectedIdFieldName].Value), wi.Id);
                    }
                }
                catch (ArgumentException)
                {
                    log(TraceLevel.Error, "Warning: Work Item " + wi.Id + " points to " + int.Parse((string)wi.Fields[ReflectedIdFieldName].Value) + " but another WI already points to this one. Ignoring...");
                }
            }
        }

        public class AddedLinks
        {
            public int RelatedLinks;
            public int ChangesetLinks;
            public int ExternalLinks;
            public List<Link> Links = new List<Link>();
        }


        public void ProcessWorkItems(string targetWorkItemsQuery)
        {
            if (_sourceTpc == null) Initialize();

            var id = FindQuery(_targetWis.Projects[TargetProject].QueryHierarchy, targetWorkItemsQuery);

            if (id == Guid.Empty)
            {
                log(TraceLevel.Error, "Unable to find query {0}", targetWorkItemsQuery);
                return;
            }

            log(TraceLevel.Info, "Found query '{0}' with id '{1}'", targetWorkItemsQuery, id);

            ProcessWorkItems(id);
        }

        public void ProcessWorkItems(Guid targetWorkItemsQueryGuid)
        {
            if (_sourceWis == null) Initialize();

            log(TraceLevel.Info, "Querying target work items...");
            var queryDef = _targetWis.GetQueryDefinition(targetWorkItemsQueryGuid);

            log(TraceLevel.Info, "Executing query " + queryDef.Name);
            log(TraceLevel.Verbose, queryDef.QueryText);
            WorkItemCollection result = _targetWis.Query(queryDef.QueryText.Replace("@project", "'" + TargetProject + "'"));
            log(TraceLevel.Info, result.Count + " Work Item(s) found");

            log(TraceLevel.Info, "Starting processing");
            foreach (WorkItem wi in result)
            {
                if (wi.Fields.Contains(ReflectedIdFieldName))
                {
                    Field reflectedWorkItemIdField = wi.Fields[ReflectedIdFieldName];
                    int relfectedWorkItemId;
                    if (int.TryParse((string)reflectedWorkItemIdField.Value, out relfectedWorkItemId))
                    {
                        WorkItem sourceWi = _sourceWis.GetWorkItem(relfectedWorkItemId);

                        log(TraceLevel.Verbose, "Target workitem[{0}] Source workitem[{1}] ExistingLinks={2} SourceLinks={3}", wi.Id, sourceWi.Id, wi.Links.Count, sourceWi.Links.Count);

                        // computes the missing links (but doesn't assign them yet to the target WI)
                        AddedLinks addedLinks = reflectLinks(sourceWi, wi);

                        // if missing links have been found
                        if (addedLinks.Links.Count > 0)
                        {
                            // try to assign the links for real in the targetWi
                            try
                            {
                                // fully reload the target WorkItem before doing any change, if not, we *may* have TF26212 errors
                                // TF26212: Team Foundation Server could not save your changes. There may be problems with the work item type definition. Try again or contact your Team Foundation Server Administrator.
                                WorkItem targetWi = _targetWis.GetWorkItem(wi.Id);

                                addedLinks.Links.ForEach(x => targetWi.Links.Add(x));
                                targetWi.Save();

                                // done, update the stats
                                NbTargetRelatedLinksAdded += addedLinks.RelatedLinks;
                                NbTargetChangesetLinksAdded += addedLinks.ChangesetLinks;
                                NbTargetExternalLinksAdded += addedLinks.ExternalLinks;
                                _processedWorkItems.Add(targetWi.Id);
                            }
                            catch (Exception ex)
                            {
                                log(TraceLevel.Error, "Exception: " + ex.Message);
                                NbSaveErrors++;
                            }
                        }
                    }
                    else
                    {
                        log(TraceLevel.Error, "Error: could not parse {0} on work item {1}", ReflectedIdFieldName, wi.Id);
                    }
                }
                else
                {
                    log(TraceLevel.Error, "Error: missing {0} field on work item {1}", ReflectedIdFieldName, wi.Id);
                }
                NbProcessedWorkItems++;
            }

        }

        /// <summary>
        /// Will create copies of missing links, but won't create them yet in TFS: only computing missing links here.
        /// </summary>
        /// <param name="sourceWi">The source Work item</param>
        /// <param name="targetWi">The target Work item: do not modify targetWi</param>
        /// <returns>An AddedLinks structure which contains stats and the copied links to be added</returns>
        private AddedLinks reflectLinks(WorkItem sourceWi, WorkItem targetWi)
        {
            NbTargetRelatedLinksFound += targetWi.RelatedLinkCount;
            int nbChangesetsLinks = targetWi.Links.OfType<ExternalLink>().Where(l => isChangeset(l)).Count();
            NbTargetExternalLinksFound += targetWi.ExternalLinkCount - nbChangesetsLinks;
            NbTargetChangesetLinksFound += nbChangesetsLinks;
            var addedLinks = new AddedLinks();

            foreach (Link link in sourceWi.Links)
            {
                Link newLink = null;

                var hyperLink = link as Hyperlink;
                if (hyperLink != null)
                {
                    if (!targetWi.Links.OfType<Hyperlink>().Any(x => x.Location == hyperLink.Location))
                    {
                        log(TraceLevel.Warning, "Warning: a missing hyperlink has been found on work item " + targetWi.Id);
                        newLink = new Hyperlink(hyperLink.Location);
                        log(TraceLevel.Verbose, "Adding Hyperlink to WI " + targetWi.Id + " pointing to " + hyperLink.Location);
                    }
                }

                // related links point to other work items, these are the ones we need to transpose
                var relatedLink = link as RelatedLink;
                if (relatedLink != null)
                {
                    NbSourceRelatedLinksFound++;
                    if (AddMissingRelatedLinks)
                    {
                        // first get what the target link related work item Id should be
                        int relatedReflectedWorkItemId;
                        if (_reflectedWorkItemIds.TryGetValue(relatedLink.RelatedWorkItemId, out relatedReflectedWorkItemId))
                        {
                            // sometimes links there are cross links whithin the same scope
                            // we must avoid to add twice the same link to avoid "link already exists" which fails the Save method later 
                            // this can happen because of the cache not reflecting what has been done previously
                            // if the link points to some work item already processed, the link has already been added, then skip it
                            if (!_processedWorkItems.Contains(relatedReflectedWorkItemId))
                            {
                                bool sameType = targetWi.Links.OfType<RelatedLink>().Any(x => x.LinkTypeEnd.ImmutableName == relatedLink.LinkTypeEnd.ImmutableName);
                                bool sameWi = targetWi.Links.OfType<RelatedLink>().Any(x => x.RelatedWorkItemId == relatedReflectedWorkItemId);

                                if (sameWi && !sameType) log(TraceLevel.Warning, "Warning: found 2 links of differents types targetting the same WI on " + targetWi.Id);
                                // check if the same link already exists
                                if (!(sameType && sameWi))
                                {
                                    WorkItemLinkTypeEnd wilte = _targetWis.WorkItemLinkTypes.LinkTypeEnds.Where(lte => lte.ImmutableName == relatedLink.LinkTypeEnd.ImmutableName).FirstOrDefault();
                                    if (wilte != null)
                                    {
                                        newLink = new RelatedLink(relatedLink.LinkTypeEnd, relatedReflectedWorkItemId);
                                        addedLinks.RelatedLinks++;
                                        log(TraceLevel.Verbose, "Adding Related link to WI " + targetWi.Id + " pointing to WI " + relatedReflectedWorkItemId);
                                    }
                                    else
                                    {
                                        log(TraceLevel.Error, "Error: link type end " + relatedLink.LinkTypeEnd.ImmutableName + " not found in target work item store");
                                    }
                                }
                            }
                            else
                            {
                                NbTargetCrossRelatedLinks++;
                            }
                        }
                        else
                        {
                            log(TraceLevel.Warning, "Warning: could not find target related WI {0} for source WI {1}", relatedLink.RelatedWorkItemId, sourceWi.Id);
                            NbMissingRelatedWorkItems++;
                        }
                    }
                }

                var externalLink = link as ExternalLink;
                if (externalLink != null)
                {
                    var externalLinkType = externalLink.ArtifactLinkType;
                    var externalLinkUri = externalLink.LinkedArtifactUri;

                    // changesets are external links, we need to distinguish them
                    bool changeset = isChangeset(externalLink);

                    if (changeset)
                    {
                        NbSourceChangesetLinksFound++;

                        var sourceChangeset = _sourceVersionControlServer.ArtifactProvider.GetChangeset(new Uri(externalLinkUri));
                        var hasSourceChangesetId = _targetVersionControlServer.GetAllCheckinNoteFieldNames()
                            .Any(m => m.Equals("SourceChangesetId", StringComparison.OrdinalIgnoreCase));

                        if (hasSourceChangesetId)
                        {
                            // Create versionspec's. Example start with changeset 1
                            VersionSpec versionFrom = VersionSpec.ParseSingleSpec("C1", null);
                            // If you want all changesets use this versionFrom:
                            // VersionSpec versionFrom = null;
                            VersionSpec versionTo = VersionSpec.Latest;

                            // Get Changesets
                            var changesets = _targetVersionControlServer.QueryHistory(
                              "$/",
                              VersionSpec.Latest,
                              0,
                              RecursionType.Full,
                              null,
                              versionFrom,
                              versionTo,
                              Int32.MaxValue,
                              false,
                              false
                              ).Cast<Changeset>();

                            var targetChangeset = changesets.FirstOrDefault(
                                m =>
                                {
                                    var fieldValue = m.CheckinNote.Values.FirstOrDefault(
                                        x => x.Name.Equals("SourceChangesetId", StringComparison.OrdinalIgnoreCase));
                                    return fieldValue != null && fieldValue.Value == sourceChangeset.ChangesetId.ToString();
                                });

                            if (targetChangeset != null)
                            {
                                externalLinkUri = targetChangeset.ArtifactUri.AbsoluteUri;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    else
                        NbSourceExternalLinksFound++;

                    if ((changeset && AddMissingChangesetsLinks) || (!changeset && AddMissingExternalLinks))
                    {
                        // Warning: use Equals for ArtifactLinkType (referenceequals or == does not work)
                        if (!targetWi.Links.OfType<ExternalLink>().Any(x => x.ArtifactLinkType.Equals(externalLinkType) && x.LinkedArtifactUri == externalLinkUri))
                        {
                            newLink = new ExternalLink(externalLinkType, externalLinkUri);
                            if (changeset)
                                addedLinks.ChangesetLinks++;
                            else
                                addedLinks.ExternalLinks++;
                            log(TraceLevel.Verbose, "Adding External link to wi " + targetWi.Id + " " + externalLinkUri);
                        }
                    }
                }


                if (newLink != null)
                {
                    // link common fields
                    // not sure about what a "locked" link is, let's not replicate them as "locked" and ignore the IsLock property
                    // IsNew : read-only
                    newLink.Comment = link.Comment;
                    addedLinks.Links.Add(newLink);

                    // cap for updates because too many links at once would result in errors during Save
                    if (addedLinks.Links.Count == MAX_WI_LINKS_TO_RESTORE)
                    {
                        log(TraceLevel.Warning, "Warning: Work Item " + targetWi.Id + " has reached max number of links to restore");
                        break;
                    }
                }
            }
            return addedLinks;
        }

        private static bool isChangeset(ExternalLink link)
        {
            return link.ArtifactLinkType.Name == "Fixed in Changeset";
        }

        #region Logging helper methods

        private void log(TraceLevel level, string message)
        {
            if (LogMessageEvent != null)
            {
                LogMessageEvent.Invoke(message, level);
            }
        }

        private void log(TraceLevel level, string format, params object[] args)
        {
            if (LogMessageEvent != null)
            {
                LogMessageEvent.Invoke(String.Format(format, args), level);
            }
        }

        #endregion

        public void Dispose()
        {
            _sourceTpc.Dispose();
            _targetTpc.Dispose();
        }
    }

}
