using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Framework.Client;
using System.Linq;

namespace TfsManager
{
    public partial class ChangesetFinder : Form
    {
        public ChangesetFinder()
        {
            InitializeComponent();
        }
        public class ChangesetItem
        {
            public string Comment { get; set; }
            public long WorkItemId { get; set; }
            public long ChangesetId { get; set; }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
                return;

            if (string.IsNullOrEmpty(textBox2.Text))
                return;
            richTextBox1.Text = string.Empty;
            var tfsUri = new Uri(textBox1.Text);

            var tpc = new TfsTeamProjectCollection(tfsUri);

            var workItemStore = (WorkItemStore)tpc.GetService(typeof(WorkItemStore));
            var configurationServer = TfsConfigurationServerFactory.GetConfigurationServer(tfsUri);
            var tpcService = configurationServer.GetService<ITeamProjectCollectionService>();

            var arr = textBox2.Text.Split(',');

            List<long> results = new List<long>();
            foreach (string s in arr)
            {
                long val;

                if (long.TryParse(s.Trim(), out val))
                {
                    results.Add(val);
                }
            }

            long[] workItemList = results.Distinct().ToArray();

            var workItems = new List<WorkItem>();

            foreach (var id in workItemList)
            {
                var wi = workItemStore.GetWorkItem((int)id);
                if (wi != null)
                {
                    workItems.Add(wi);
                }
            }

            var list = new List<ChangesetItem>();

            foreach (var workItem in workItems)
            {
                var cs = GetChangesetList(workItem, workItemStore);

                if (cs != null && cs.Count > 0)
                {
                    list.AddRange(cs);
                }
            }

            list = list.OrderBy(x => x.ChangesetId).Distinct().ToList();
            richTextBox1.Text += 0 + "\t" + "ChangesetId" + "\t" + "WorkItemId" + "\t" + "Comment" + System.Environment.NewLine;
            for (var i = 0; i < list.Count; i++)
            {
                richTextBox1.Text += i + 1 + "\t" + list[i].ChangesetId + "\t" + list[i].WorkItemId + "\t" + list[i].Comment + System.Environment.NewLine;
            }
        }

        private List<ChangesetItem> GetChangesetList(WorkItem workItem, WorkItemStore ws)
        {
            var list = new List<ChangesetItem>();
            foreach (var link in workItem.Links)
            {
                if (link == null)
                    continue;

                if (link is ExternalLink)
                {
                    string externalLink = ((ExternalLink)link).LinkedArtifactUri;
                    var artifact = LinkingUtilities.DecodeUri(externalLink);

                    if (artifact.ArtifactType == "Changeset")
                    {
                        long csId = 0;

                        if (long.TryParse(artifact.ToolSpecificId, out csId))
                        {
                            var csItem = new ChangesetItem();
                            csItem.ChangesetId = csId;
                            csItem.WorkItemId = workItem.Id;
                            csItem.Comment = "";
                            list.Add(csItem);
                        }
                    }
                }
                else if (link is RelatedLink)
                {

                    var rLink = ((RelatedLink)link);

                    if (rLink.LinkTypeEnd.Name == "Child")
                    {
                        var id = ((RelatedLink)rLink).RelatedWorkItemId;

                        var childWorkItem = ws.GetWorkItem(id);

                        if (childWorkItem != null)
                        {
                            var childCs = GetChangesetList(childWorkItem, ws);

                            if (childCs != null && childCs.Count > 0)
                            {
                                list.AddRange(childCs);
                            }
                        }
                    }

                }
            }
            return list;
        }

        private WorkItemCollection GetWorkItemById(WorkItemStore workItemStore, long workItemId)
        {
            var queryResults = workItemStore.Query(
            string.Format("Select [Title] " +
            "From WorkItems " +
            "Where [ID] = '{0}' ", workItemId));
            return queryResults;
        }

    }
}
