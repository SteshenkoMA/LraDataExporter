using System;
using System.Windows.Forms;
using Analysis.Api;
using Analysis.ApiLib;
using Analysis.Api.Notifications;
using System.Diagnostics;

namespace AnalysisAPISample
{
    partial class MainForm
    {
        /// <summary>
        /// Event handler for NotifyStatusUpdate events with NotificationProgressData argument
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">NotificationProgressData argument</param>
        private void APIProgressEvent(object sender, NotificationProgressData e)
        {
            toolStripProgressBar1.Visible = true;
            if (e.Progress == -1)
            {
                e.CanContinue = true; 
            }
            else
            {
                if (e.MaxProgress > 0)
                    toolStripProgressBar1.Maximum = e.MaxProgress;
                toolStripProgressBar1.Value = e.Progress;
                statusStrip1.Refresh();
                toolStripStatusLabel1.Text = e.Task;
            }
        }

        /// <summary>
        /// Event handler for NotifyStatusUpdate events with NotificationStatusData argument
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">NotificationStatusData argument</param>
        void APIStatusUpdate(object sender, NotificationStatusData e)
        {
            toolStripStatusLabel1.Text = e.Task;    
        }

        /// <summary>
        /// Event handler for CreateHtmlReport.After event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">NotificationsHTMLReportData argument</param>
        private void APICreateHtmlReport(object sender, NotificationsHTMLReportData e)
        {
            UseWaitCursor = false;
 
            //open created report
            Process.Start("iexplore.exe", e.ReportFileName); 
        }

        /// <summary>
        /// Event handler for *.After events
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void APIAfterEvent(object sender, EventArgs e)
        {
            UseWaitCursor = false;
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabel1.Text = String.Empty;
        }

        /// <summary>
        /// Event handler for *.Before events
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void APIBeforeEvent(object sender, EventArgs e)
        {
            UseWaitCursor = true;
        }
    }

    /// <summary>
    /// EventArgs descendant used for internal communication
    /// </summary>
    class GraphEventArgs : EventArgs
    {

        //graph name
        private string _graph;
        //series name
        private string _series;

        public string graph
        {
            get
            {
                return _graph;
            }
            set
            {
                _graph = value;
            }
        }

        public string series
        {
            get
            {
                return _series;
            }
            set
            {
                _series = value;
            }
        }

        public GraphEventArgs(string graph, string series)
        {
            this._graph = graph;
            this._series = series;
        }
    } 
}