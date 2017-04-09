using System;
using System.Drawing;
using System.Windows.Forms;
using Analysis.Api;
using Analysis.ApiLib;
using Analysis.Api.Notifications;
using Analysis.Api.Export;
using Analysis.Api.Dictionaries;
using Analysis.ApiLib.Dimensions;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using Analysis.ApiLib.Sla;

namespace AnalysisAPISample
{
    partial class MainForm
    {
        /// <summary>
        /// Initializes event handlers.
        /// </summary>
        private void InitializeEvents()
        {
            NotificationSessionEvents events = _analysisApi.Session.Events;
            Exporters export = _analysisApi.Exporters;

            //events of the Session object
            events.SessionCreate.Before += new EventHandler<EventArgs>(APIBeforeEvent);
            events.SessionCreate.After += new EventHandler<EventArgs>(APIAfterEvent);
            events.SessionOpen.Before += new EventHandler<EventArgs>(APIBeforeEvent);
            events.SessionOpen.After += new EventHandler<EventArgs>(APIAfterEvent);
            events.RunLoadGraphList.Before += new EventHandler<EventArgs>(APIBeforeEvent);
            events.RunLoadGraphList.After += new EventHandler<EventArgs>(APIAfterEvent); 
            events.GraphApplyFilterAndGroupBy.Before += new EventHandler<EventArgs>(APIBeforeEvent);
            events.GraphApplyFilterAndGroupBy.After += new EventHandler<EventArgs>(APIAfterEvent); 
            events.CreateHtmlReport.After += new EventHandler<NotificationsHTMLReportData>(APICreateHtmlReport);
            events.CreateHtmlReport.Before += new EventHandler<NotificationsHTMLReportData>(APIBeforeEvent);
            events.SessionCreate.NotifyStatusUpdate += new EventHandler<NotificationProgressData>(APIProgressEvent);
            events.RunLoadGraphList.NotifyStatusUpdate += new EventHandler<NotificationProgressData>(APIProgressEvent);

            //events of Exporters object
            export.CSV.Events.Graph.Before += new EventHandler<EventArgs>(APIBeforeEvent);
            export.CSV.Events.Graph.After += new EventHandler<EventArgs>(APIAfterEvent); 
            export.CSV.Events.Dictionary.NotifyStatusUpdate += new EventHandler<NotificationStatusData>(APIStatusUpdate);
            export.CSV.Events.Series.NotifyStatusUpdate += new EventHandler<NotificationStatusData>(APIStatusUpdate);
            export.CSV.Events.Graph.NotifyStatusUpdate += new EventHandler<NotificationProgressData>(APIProgressEvent);
        }

        /// <summary>
        /// Opens a session using the database file.
        /// </summary>
        /// <param name="file_name">The absolute path name of the session database file. 
        /// Database files have an lra extension.</param>
        /// <returns>true if session opened successfully</returns>
        private bool OpenSession(string file_name)
        {
            bool result = _analysisApi.Session.Open(file_name);
            if (result)
            {
                //get references for easy access
               
                _currentSession = _analysisApi.Session;
                _currentRun = _currentSession.Runs[_usedRun];
            }
            else
            {
                //show error message
                MessageBox.Show(_analysisApi.RunTimeErrors.LastErrorMessage);
            }
            return result;
        }

        /// <summary>
        /// Inserts the list of graphs and their series to the TreeView.
        /// </summary>
        private void GetGraphsFromRun()
        {
            GraphView.Nodes.Clear();
            GraphView.BeginUpdate();

            //reference for easy access
            GraphsList graphs = _currentRun.Graphs;
            try
            {
                //iterate through all graphs
                for (int i = 0; i < graphs.Count; i++)
                {
                    //add a graph to TreeVeiw
                    GraphView.Nodes.Add(graphs[i].Name.Name, graphs[i].Name.Name, 1);
                    try
                    {
                        //if there are no series in graph - mark it as an empty
                        if (graphs[i].Series.Count == 0)
                            GraphView.Nodes[i].ImageIndex = _EmptyGraph;

                        //iterate through all series of current graph
                        foreach (Series s in graphs[i].Series)
                        {
                            //add series as a child of a graph
                            GraphView.Nodes[i].Nodes.Add(s.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        //mark problematic graph as empty
                        GraphView.Nodes[i].ImageIndex = _EmptyGraph;
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            finally
            {
                GraphView.EndUpdate();
            }

            //uncheck show all graphs menu item
            showAllGraphsToolStripMenuItem.Checked = false;
        }

        /// <summary>
        /// Creates a session using the specified result file and saves the new session to a file.
        /// </summary>
        /// <param name="session_name">The absolute path name of the new session file. 
        /// The file has an lra extension.</param>
        /// <param name="res_name">The absolute path name of the existing results file. 
        /// The file has an lrr extension. This result is added to the session. </param>
        /// <returns>true if session created successfully</returns>
        private bool CreateSession(string session_name, string res_name)
        {
            bool result = _analysisApi.Session.Create(session_name, res_name);
            if (result)
            {
                //get references for easy access
                _currentSession = _analysisApi.Session;
                _currentRun = _currentSession.Runs[_usedRun];
            }
            else
            {
                //set all to null
                _currentSession = null;
                _currentRun = null;
                //show error message
                MessageBox.Show(_analysisApi.RunTimeErrors.LastErrorMessage);

            }
            return result;
        }

        /// <summary>
        /// Closes the previously opened session.
        /// </summary>
        /// <returns>true if session closed successfully</returns>
        private bool CloseSession()
        {
            bool result = _analysisApi.Session.Close();
            _lastGraph = String.Empty;
            _currentSession = null;
            _currentRun = null;
            return result;
        }
        
        /// <summary>
        /// Creates an HTML report and saves it to a file.
        /// </summary>
        /// <param name="file_name">report file's pathname</param>
        /// <returns>true if report created and saved successfully</returns>
        private bool CreateHTMLReport(string file_name)
        {
            if (_currentSession.IsOpenedOrCreated == false)
                return false;

            UseWaitCursor = true;
            //get a reference to ReportsMaker
            HtmlReportMaker maker = _currentSession.CreateHtmlReportMaker();
            //create a default report
            bool result = maker.CreateDefaultHtmlReport(file_name, ApiBrowserType.IE);
            
            return result;
        }

        /// <summary>
        /// Fills the ListView with information on the current session and run.
        /// </summary>
        private void GetRunInformation()
        {
            DateTime st_time = new DateTime(1970, 1, 1, 0, 0, 0);
            st_time = st_time.AddSeconds(_currentRun.StartTime);
            DateTime end_time = new DateTime(1970, 1, 1, 0, 0, 0);
            end_time = end_time.AddSeconds(_currentRun.EndTime);
            TimeSpan duration = end_time - st_time;

          //  Console.WriteLine(st_time +" "+ _currentRun.StartTime);
          //  Console.WriteLine(end_time);
         //   Console.WriteLine(duration);

            infoView.Items.Add("Scenario name").SubItems.Add(_currentRun.Scenario);
            infoView.Items.Add("Result name").SubItems.Add(_currentRun.Name);
            infoView.Items.Add("Duration").SubItems.Add(duration.ToString());
            infoView.Items.Add("Graphs count in session").SubItems.Add(_currentRun.Graphs.Count.ToString());
            infoView.Items.Add("Total graphs count").SubItems.Add(
                ApiGlobal.GetInstance().GraphNames.Count.ToString());
            infoView.Items.Add("Session file name").SubItems.Add(
                _currentSession.Name);
        }

        /// <summary>
        /// Converts the time in seconds to a string.
        /// </summary>
        /// <param name="d_time">time in seconds</param>
        /// <returns>string representation of the specified time</returns>
        private string DecodeTime(double d_time)
        {
            int hours, minutes, seconds;
            int time = (int)Math.Ceiling(d_time);

            hours = time / 3600;
            time = time - (hours * 3600);
            minutes = time / 60;
            seconds = time % 60;

            //result string 00:00:00
            return hours.ToString("00") + ":" + minutes.ToString("00") + ":" + seconds.ToString("00");
        }

        /// <summary>
        /// Extracts data for the specified graph and populates the ListView with the data.
        /// </summary>
        /// <param name="graph_name">name of graph</param>
        /// <param name="series_name">name of series to be extracted</param>
        /// <param name="g_data">reference to a GraphDataForm object</param>
        /// <param name="use_iterator">specify whether to use an iterator when extracting data</param>
        public void OpenGraphData(string graph_name, string series_name, GraphDataForm g_data, 
            bool use_iterator)
        {
            if (_currentSession.IsOpenedOrCreated == false)
                return;

            Graph current_graph = null;
            if (null == g_data)
                g_data = new GraphDataForm();
            GraphsList g_list = _currentRun.Graphs;
            Series current_series = null;
            current_graph = g_list[graph_name];
            current_series = current_graph.Series[series_name];
            
            //PointsDefinition class used to get names of values
            for (int i = 0; i < current_series.PointsDefinition.Count; i++)
                g_data.AddColumnToListView(current_series.PointsDefinition[i].Name);

            //add all graph series to ComboBox
            foreach (Series s in current_graph.Series)
            {
                g_data.AddSeries(s.Name);
            }
            //select current series
            g_data.SelectSeries(current_series.Name);

            g_data.Text = current_graph.Name.DisplayName;
            g_data.CurrentGraphName = current_graph.Name.Name;

            //fill ListView with data of current series
            FillDataIntoListView(current_series, g_data, current_graph.HasTimeAxis, use_iterator);

            _lastGraph = current_graph.Name.Name;

            //show the form
            if (false == g_data.Visible)
                g_data.ShowDialog();
        }

        /// <summary>
        /// Get a string representation for statistics
        /// </summary>
        /// <param name="stat">reference to a statistics object</param>
        /// <param name="function">statistics function type</param>
        /// <returns>string representation of specified statistics</returns>
        private string GetStatistics(CommonSeriesStatistics stat, StatisticsFunctionKind function)
        {
            string result = String.Empty;

            SeriesGraphStatistics s_gs = null;
            SeriesRawStatistics s_rs = null;

            //check statistics type
            if (stat is SeriesRawStatistics)
            {
                s_rs = (SeriesRawStatistics)stat;
 
                //if statistics are available, store in variable
                if (s_rs.IsFunctionAvailable(function))
                {
                    switch(function)
                    {
                        //statistics is calculated with 3 digits precision
                        case StatisticsFunctionKind.Average:
                            result = s_rs.Average.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Count:
                            result = s_rs.Count.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Maximum:
                            result = s_rs.Maximum.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Minimum:
                            result = s_rs.Minimum.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.StdDeviation:
                            result = s_rs.StdDeviation.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Sum:
                            result = s_rs.Sum.ToString("0.000");
                            break;
                        default:
                            result = "Unknown function";
                            break;
                    }
                }
                else
                    //statistics not available
                    result = "N/A";
            }

            //check statistics type
            if (stat is SeriesGraphStatistics)
            {
                s_gs = (SeriesGraphStatistics)stat;

                //if statistics are available, store in variable
                if (s_gs.IsFunctionAvailable(function))
                {
                    switch (function)
                    {
                        case StatisticsFunctionKind.Average:
                            result = s_gs.Average.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Maximum:
                            result = s_gs.Maximum.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Minimum:
                            result = s_gs.Minimum.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.Median:
                            result = s_gs.Median.ToString("0.000");
                            break;
                        case StatisticsFunctionKind.StdDeviation:
                            result = s_gs.StdDeviation.ToString("0.000");
                            break;
                        default:
                            result = "Unknown function";
                            break;
                    }
                }
                else
                    //statistics not available
                    result = "N/A";
            }

            return result;
        }

        /// <summary>
        /// Populates the ListView with all statistics.
        /// </summary>
        /// <param name="series">reference to a series object</param>
        /// <param name="granularity">granularity of current graph</param>
        private void FillGraphStatistics(Series series, int granularity)
        {
            if (null == series)
                return;

            graphStatistics.Items.Clear();
            SeriesGraphStatistics stat = series.GraphStatistics;

            graphStatistics.Items.Add("Average").SubItems.Add(GetStatistics(stat, StatisticsFunctionKind.Average));
            graphStatistics.Items.Add("Minimum").SubItems.Add(GetStatistics(stat, StatisticsFunctionKind.Minimum));
            graphStatistics.Items.Add("Maximum").SubItems.Add(GetStatistics(stat, StatisticsFunctionKind.Maximum));
            graphStatistics.Items.Add("Median").SubItems.Add(GetStatistics(stat, StatisticsFunctionKind.Median));
            graphStatistics.Items.Add("Std. Deviation").SubItems.Add(GetStatistics(stat, StatisticsFunctionKind.StdDeviation));
            textBoxGranularity.Text = granularity.ToString();

            rawStatistics.Items.Clear();
            SeriesRawStatistics raw_stat = series.RawStatistics;
 
            rawStatistics.Items.Add("Average").SubItems.Add(GetStatistics(raw_stat, StatisticsFunctionKind.Average));
            rawStatistics.Items.Add("Minimum").SubItems.Add(GetStatistics(raw_stat, StatisticsFunctionKind.Minimum));
            rawStatistics.Items.Add("Maximum").SubItems.Add(GetStatistics(raw_stat, StatisticsFunctionKind.Maximum));
            rawStatistics.Items.Add("Count").SubItems.Add(GetStatistics(raw_stat, StatisticsFunctionKind.Count));
            rawStatistics.Items.Add("Std. Deviation").SubItems.Add(GetStatistics(raw_stat, StatisticsFunctionKind.StdDeviation));
            rawStatistics.Items.Add("Sum").SubItems.Add(GetStatistics(raw_stat, StatisticsFunctionKind.Sum));
        }

        /// <summary>
        /// Saves graph to a CSV file. 
        /// </summary>
        /// <param name="graph">reference to a graph object</param>
        /// <param name="file_name">output file name</param>
        private void ExportGraphToFile(Graph graph, string file_name)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;
            ExportToCSVFile(graph, file_name);
        }

        /// <summary>
        /// Saves dictionary to a CSV file.
        /// </summary>
        /// <param name="dict">reference to a dictionary object</param>
        /// <param name="file_name">output file name</param>
        private void ExportDictionaryToFile(IBaseDictionary dict, string file_name)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;
            ExportToCSVFile(dict, file_name);
        }

        /// <summary>
        /// Save series to a CSV file.
        /// </summary>
        /// <param name="series">reference to a dictionary object</param>
        /// <param name="file_name">output file name</param>
        private void ExportSeriesToFile(Series series, string file_name)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;
            ExportToCSVFile(series, file_name);
        }

        /// <summary>
        /// Saves an object to a CSV file.
        /// </summary>
        /// <remarks>The export is performed according to the object's type.</remarks>
        /// <param name="obj_save">reference to an object to save</param>
        /// <param name="file_name">output file name</param>
        private void ExportToCSVFile(object obj_save, string file_name)
        {
            try
            {
                //check object's type and then export it to file
                if (obj_save is Graph)
                    _analysisApi.Exporters.CSV.ExportGraph((Graph)obj_save, file_name);
                if (obj_save is IBaseDictionary)
                    _analysisApi.Exporters.CSV.ExportDictionary((IBaseDictionary)obj_save, file_name);
                if (obj_save is Series)
                    _analysisApi.Exporters.CSV.ExportSeries((Series)obj_save, file_name);
            }
            catch(Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Gets data for the specified series and populates the ListView.
        /// </summary>
        /// <param name="s">reference to a series object</param>
        /// <param name="g_data">reference to a GraphDataForm object</param>
        /// <param name="hastime">specify whether series has time axis</param>
        /// <param name="use_iterator">specify whether to use iterator when extracting data or not</param>
        private void FillDataIntoListView(Series s, GraphDataForm g_data, bool hastime, 
            bool use_iterator)
        {
            if (!use_iterator)
            {
                //get data using Points collection
                GetDataPoints(s, g_data, hastime);
            }
            else
            {
                //get data using iterator
                GetDataWithIterator(s, g_data, hastime);
            }
        }

        /// <summary>
        /// Gets data from a series using the Points collection.
        /// </summary>
        /// <param name="s">reference to a series object</param>
        /// <param name="g_data">reference to a GraphDataForm object</param>
        /// <param name="hastime">specify whether series has time axis</param>
        private void GetDataPoints(Series s, GraphDataForm g_data, bool hastime)
        {
            string time = null;

            //iterate through all points in series
            //If the collection is empty, use an iterator
            foreach (SeriesPoint point in s)
            {
                if (hastime)
                {
                    //decode time to hh:mm:ss
                    time = DecodeTime(point.RelativeTime);
                }
                else
                {
                    //series doesn't have a time axis
                    time = point.RelativeTime.ToString();
                }

                //create an array of points
                string[] columns = new string[point.Values.Count];
                for (int i = 0; i < point.Values.Count; i++)
                    columns[i] = point.Values[i].Value.ToString("0.00000000");//all points are calculated with 8 digits precision

                //assign created array to ListView
                g_data.ListBeginUpdate();
                g_data.FillFirstColumn(time).SubItems.AddRange(columns);
                g_data.ListEndUpdate();
            }

            /* Use the following code if you need to fire a SeriesLoadPoints event
             * 
             *  string time = null;
             * 
             *  for (int i = 0; i < s.Points.Count; i++)
             *  {
             *      if (hastime)
             *      {
             *          //decode time to hh:mm:ss
             *          time = DecodeTime(point.RelativeTime);
             *      }
             *      else
             *      {
             *          //series doesn't have a time axis
             *          time = s.Points[i].RelativeTime.ToString();
             *      }
             *
             *      //create an array with all points
             *      string[] columns = new string[s.Points[i].Values.Count];
             *      for (int y = 0; y < s.Points[i].Values.Count; y++)
             *          columns[y] = s.Points[i].Values[y].Value.ToString("0.00000000");
             *
             *      //assign created array to ListView
             *      g_data.ListBeginUpdate();
             *      g_data.FillFirstColumn(time).SubItems.AddRange(columns);
             *      g_data.ListEndUpdate();
             *  }
             * */
        }

        /// <summary>
        /// Gets data from a series using an iterator.
        /// </summary>
        /// <param name="s">reference to a series object</param>
        /// <param name="g_data">reference to a GraphDataForm object</param>
        /// <param name="hastime">specify whether series has time axis</param>
        private void GetDataWithIterator(Series s, GraphDataForm g_data, bool hastime)
        {
            //check if iterator can be used
            if (!s.CanUsePointsIterator())
            {
                MessageBox.Show("Iterator cannot be used");
                return;
            }
            //get iterator
            SeriesPointIterator pi = s.GetPointsIterator();
            string time = null;

            //iterate through all points using iterator
            foreach (SeriesPoint pt in pi)
            {
                if (hastime )
                {
                    //decode time to hh:mm:ss
                    time = DecodeTime(pt.RelativeTime);
                }
                else
                {
                    //series doesn't have a time axis
                    time = pt.RelativeTime.ToString();
                }

                //create an array with all points
                string[] columns = new string[pt.Values.Count];
                for (int i = 0; i < pt.Values.Count; i++)
                    columns[i] = pt.Values[i].Value.ToString("0.00000000");

                //assign created array to ListView
                g_data.ListBeginUpdate();
                g_data.FillFirstColumn(time).SubItems.AddRange(columns);
                g_data.ListEndUpdate();
            }  
        }

        /// <summary>
        /// Gets Filter and GroupBy items from a Graph.
        /// </summary>
        /// <param name="graph">reference to a graph object</param>
        private void GetFiltersAndGroupBy(Graph graph)
        {
            GetGroupByItems(graph);
            GetFilterItems(graph);
        }

        /// <summary>
        /// Gets all GroupBy items from a graph and populates a list box.
        /// </summary>
        /// <param name="graph">reference to a graph object</param>
        private void GetGroupByItems(Graph graph)
        {
            checkedListBox1.Items.Clear();
            foreach(GroupByItem gb in graph.GroupBy)
            {
                checkedListBox1.Items.Add(gb.Name, gb.IsActive);
            }
        }

        /// <summary>
        /// Gets the state of the global filter and populates the form controls.
        /// </summary>
        private void GetGlobalFilter()
        {
            //get reference to a GlobalFilter object
            ApiGlobalFilter filter = _currentSession.GlobalFilter;

            checkBox1.Checked = filter.IncludeThinkTime;
            //There is always one and only one value in ScenarioElapsedTime
            textBox2.Text = filter.ScenarioElapsedTime.AvailableValues.ContinuousValues[0].Value.Min.ToString();
            textBox3.Text = filter.ScenarioElapsedTime.AvailableValues.ContinuousValues[0].Value.Max.ToString();
        }

        /// <summary>
        /// Gets the filter items from a graph and populates the ListView.
        /// </summary>
        /// <param name="graph">reference to a graph object</param>
        private void GetFilterItems(Graph graph)
        {
            filterView.Items.Clear();

            //iterate through all filter items in filter
            foreach (FilterItem fi in graph.Filter)
            {
                //create an array of the filter's properties
                string[] columns = new string[3];
                //filter name
                columns[0] = fi.Name; 
                //operator
                columns[1] = fi.ConditionalOperator.ToString();

                //check whether values are continuous or discrete
                bool cont = (fi.FilterValues.ValuesKind != FilterValuesKind.Discrete);

                //fill array with filter values
                for (int i = 0;
                    i < (false == cont ? fi.FilterValues.DiscreteValues.Count :
                        fi.FilterValues.ContinuousValues.Count); i++)
                {
                    if(!cont)
                        columns[2] += fi.FilterValues.DiscreteValues[i] + "; ";
                    else
                        columns[2] += fi.FilterValues.ContinuousValues[i].Value.Min.ToString() + 
                            ".." + fi.FilterValues.ContinuousValues[i].Value.Max.ToString() + "; ";
                }

                //assign the array to the ListView
                filterView.Items.Add(fi.IsActive.ToString()).SubItems.AddRange(columns);
            }
        }

        /// <summary>
        /// Sets GroupBy for specified graph
        /// </summary>
        /// <param name="group_by">name of GroupBy item</param>
        /// <param name="isActive">specify whether item is active</param>
        private void SetGroupByForCurrentGraph(string group_by, bool isActive)
        {
            if (_lastGraph == String.Empty)
                return; 
            Graph graph = _analysisApi.Session.Runs[0].Graphs[_lastGraph];
            
            graph.GroupBy.IsActive = true;

            //set a state of an item
            graph.GroupBy[group_by].IsActive = isActive;
        }

        /// <summary>
        /// Apples Filter and GroupBy to the specified Graph.
        /// </summary>
        /// <param name="gb_graph">name of the graph</param>
        private void ApplyFilterAndGroupBy(string gb_graph)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;
            //get a reference to a graph using its name
            Graph graph = _currentRun.Graphs[gb_graph];

            //apply filter and GroupBy
            graph.ApplyFilterAndGroupBy(); 
        }

        /// <summary>
        /// Applies the GlobalFilter.
        /// </summary>
        /// <param name="think_time">include think time</param>
        /// <param name="min">scenario elapsed time minimum</param>
        /// <param name="max">scenario elapsed time maximum</param>
        private void ApplyGlobalFilter(bool think_time, double min, double max)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;

            //check if the values are correct
            if (_currentSession.GlobalFilter.ScenarioElapsedTime.AvailableValues.CheckContinuousValue(String.Empty, min, max))
            {
                _currentSession.GlobalFilter.IncludeThinkTime = think_time;
                _currentSession.GlobalFilter.ScenarioElapsedTime.ClearValues();

                _currentSession.GlobalFilter.ScenarioElapsedTime.AddContinuousValue(min, max);
               

                //apply global filter
                _currentSession.GlobalFilter.ApplyFilter();
                                   
            }
            else
                MessageBox.Show("Value out of range");
        }

        /// <summary>
        /// Sets the granularity for the specified graph
        /// </summary>
        /// <param name="gb_graph">graph name</param>
        /// <param name="granularity">new granularity</param>
        private void SetGranularityToGraph(string gb_graph, int granularity)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;

            //get a reference to a graph using its name
            Graph graph = _currentRun.Graphs[gb_graph];

            //Granularity must be >= 1 or we get an exception
            if (granularity > 0)
                graph.Granularity = granularity;
            else
                MessageBox.Show("Granularity must be >= 1");
        }

        /// <summary>
        /// Clears all filter values. Call this function before applying new filter values.
        /// </summary>
        /// <param name="filter_name">name of the filter</param>
        private void ClearFilter(string filter_name)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;

            //get a reference to a graph using its name
            Graph graph = _currentRun.Graphs[_lastGraph];
            //get a reference to a filter using its name
            FilterItem fi = graph.Filter[filter_name];
            //clear filter values
            fi.ClearValues();
        }

        /// <summary>
        /// Checks whether a discrete value is valid.
        /// </summary>
        /// <param name="value">checked value</param>
        /// <param name="filter_name">name of the filter</param>
        /// <returns>true if value is valid</returns>
        public bool CheckDiscreteValue(string value, string filter_name)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return false;
            bool result = false;

            //get a reference to a graph using its name
            Graph graph = _currentRun.Graphs[_lastGraph];

            //get a reference to a filter using its name
            FilterItem fi = graph.Filter[filter_name];

            //check the value
            result = fi.AvailableValues.CheckDiscreteValue(value);

            return result;
        }

        /// <summary>
        /// Checks whether continuous values are valid.
        /// </summary>
        /// <param name="values">array of continuous values</param>
        /// <param name="filter_name">name of the filter</param>
        /// <returns>true if values are valid</returns>
        public bool CheckContinuousValues(SimpleFilter[] values, string filter_name)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return false;
            bool result = true;

            //get a reference to a graph using its name
            Graph graph = _analysisApi.Session.Runs[0].Graphs[_lastGraph];

            //get a reference to a filter using its name
            FilterItem fi = graph.Filter[filter_name];

            //iterate through all values in SimpleFilter array
            foreach (SimpleFilter filter in values)
            {
                //check the value
                if (fi.AvailableValues.CheckContinuousValue(filter.Name, filter.Min, filter.Max))
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
            }

            return result;
        }

        /// <summary>
        /// Adds continuous values to a filter.
        /// </summary>
        /// <param name="values">array of continuous values</param>
        /// <param name="filter_name">name of the filter</param>
        /// <param name="active">filter's state</param>
        private void AddContinuousValues(SimpleFilter[] values, string filter_name, bool active)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;

            //get a reference to a graph using its name
            Graph graph = _analysisApi.Session.Runs[0].Graphs[_lastGraph];

            //get a reference to a filter using its name
            FilterItem fi = graph.Filter[filter_name];

            //set isActive property
            fi.IsActive = active;

            //iterate through all values in SimpleFilter array
            foreach (SimpleFilter filter in values)
            {
                //add current value to filter
                fi.AddContinuousValue(filter.Name, filter.Min, filter.Max);
            }
        }

        /// <summary>
        /// Adds a discrete value to a filter.
        /// </summary>
        /// <param name="value">discrete value</param>
        /// <param name="filter_name">name of the filter</param>
        /// <param name="comparison">conditional operator</param>
        /// <param name="active">filter's state</param>
        private void AddDiscreteValue(string value, string filter_name, string comparison, bool active)
        {
            if (!_currentSession.IsOpenedOrCreated)
                return;

            //get a reference to a graph using its name
            Graph graph = _currentRun.Graphs[_lastGraph];

            //get a reference to a filter using its name
            FilterItem fi = graph.Filter[filter_name];

            //set conditional operator
            if (comparison == "=")
                fi.ConditionalOperator = FilterItem.FilterItemConditionalOperator.Equal;
            else
                fi.ConditionalOperator = FilterItem.FilterItemConditionalOperator.NotEqual;

            //set isActive property
            fi.IsActive = active;
                
            //add current value to filter
            fi.AddDiscreteValue(value);
        }

        /// <summary>
        /// Enumerate properties of the specified graph and populate ListView.
        /// </summary>
        /// <param name="graph">reference to a graph object</param>
        public void EnumProperties(Graph graph)
        {
            graphInfoView.Items.Clear();

            //get a reference to graphs type
            Type t = graph.GetType();
            //get properties
            PropertyInfo[] pi = t.GetProperties();

            //iterate through all properties
            foreach (PropertyInfo p in pi)
            {
                //add values to ListView
                if (p.Name == "HasTimeAxis" || p.Name == "Granularity")
                    graphInfoView.Items.Add(p.Name).SubItems.Add(p.GetValue(graph, null).ToString());
            }

            //get a reference to Graph.Name type
            t = graph.Name.GetType();
            //get properties
            pi = t.GetProperties();

            //iterate through all properties
            foreach (PropertyInfo p in pi)
            {
                //add values to ListView
                graphInfoView.Items.Add(p.Name).SubItems.Add(p.GetValue(graph.Name, null).ToString());
            }

        }

        /// <summary>
        /// Loads database options.
        /// </summary>
        private void LoadOptions()
        {
            //get a reference to Options object
            Options options = _analysisApi.Options;

            //check which database engine is used
            switch (options.Database.DatabaseEngine)
            {
                case DatabaseSettings.DatabaseEngines.Access2000:
                    optionsComboBox.SelectedIndex = 0;
                    break;
                case DatabaseSettings.DatabaseEngines.SqlServer2000:
                    optionsComboBox.SelectedIndex = 1;
                    break;
                    //   case DatabaseSettings.DatabaseEngines.SqLite:
                    //      optionsComboBox.SelectedIndex = 2;
                    //    break;
            }

            //get other database options used in MSSQL mode
            serverName.Text = options.Database.ServerName;
            userName.Text = options.Database.UserName;
            Password.Text = options.Database.Password;
            logicalStorage.Text = options.Database.LogicalStorageLocation;
            physicalStorage.Text = options.Database.PhysicalStorageLocation;
            windowsSecurity.Checked = options.Database.IsWindowsIntegratedSecurityUsed;
        }

        /// <summary>
        /// Saves options to file system or applies options to LrAnalysis object
        /// </summary>
        /// <param name="_save">True to save, false to apply to LRAnalysis object</param>
        private void SaveOptions(bool save)
        {
            //get a reference to Options object
            Options options = _analysisApi.Options;

            //check which database engine is used
            switch (optionsComboBox.SelectedIndex)
            {
                case 0:
                    options.Database.DatabaseEngine = DatabaseSettings.DatabaseEngines.Access2000;
                    break;
                case 1:
                    options.Database.DatabaseEngine = DatabaseSettings.DatabaseEngines.SqlServer2000;
                    break;
                    //    case 2:
                    //         options.Database.DatabaseEngine = DatabaseSettings.DatabaseEngines.SqLite;
                    //       break;
            }

            //get other database options used in MSSQL mode
            options.Database.ServerName = serverName.Text;
            options.Database.UserName = userName.Text;
            options.Database.Password = Password.Text;
            options.Database.LogicalStorageLocation = logicalStorage.Text;
            options.Database.PhysicalStorageLocation = physicalStorage.Text;
            options.Database.IsWindowsIntegratedSecurityUsed = windowsSecurity.Checked;

            //save all options to the file system
            //if not saved, options will be available only for lifetime of current LrAnalysis object.
            if (save)
                options.Save();
        }

        /// <summary>
        /// Populates SLA data into ListView
        /// </summary>
        private void GetSLAData()
        {
            if (null == _slaResult)
                return;

            //check whether SLA data exists
            if (!_slaResult.IsEmpty)
            {
                slaListView.Items.Clear();

                slaListView.Items.Add("Whole run SLA").Font = new Font(slaListView.Font, FontStyle.Bold);

                //get "whole run" SLA
                foreach (SlaWholeRunRuleResult r in _slaResult.WholeRunRules)
                {
                    ListViewItem item = slaListView.Items.Add(r.Measurement.ToString());
                    item.SubItems.Add(r.Status.ToString());
                    item.SubItems.Add(r.ActualValue.ToString());
                    item.SubItems.Add(r.GoalValue.ToString());
                    item.BackColor = Color.LightGray;
                }

                slaListView.Items.Add("");
                slaListView.Items.Add("Time range based SLA").Font = new Font(slaListView.Font, FontStyle.Bold);
                
                //get time range based SLA
                foreach (SlaTimeRangeRuleResult r in _slaResult.TimeRangeRules)
                {
                    ProcessTransactionSLA(r);
                }

                slaListView.Items.Add("");
                slaListView.Items.Add("Transaction related SLA").Font = new Font(slaListView.Font, FontStyle.Bold);

                //get transaction related SLA (average)
                for (int i = 0; i < _slaResult.TransactionRules.TimeRangeRules.Count; i++)
                {
                    SlaTransactionTimeRangeRuleResult r = _slaResult.TransactionRules.TimeRangeRules[i];
                    ProcessTransactionSLA(r);
                }

                slaListView.Items.Add("");
                slaListView.Items.Add("Percentile related SLA").Font = new Font(slaListView.Font, FontStyle.Bold);

                //get transaction related SLA (percentile)
                for (int i = 0; i < _slaResult.TransactionRules.PercentileRules.Count; i++)
                {
                    SlaPercentileRuleResult r = _slaResult.TransactionRules.PercentileRules[i];
                    ListViewItem item = new ListViewItem();
                    slaListView.Items.Add("");
                    ListViewItem pItem = slaListView.Items.Add(r.Measurement + " (" + r.TransactionName + ")");
                    pItem.Font = new Font(slaListView.Font, FontStyle.Bold);
                    SetColorByStatus(pItem, r.Status);

                    item = slaListView.Items.Add(r.Percentage.ToString() + "%");
                    item.BackColor = Color.LightGray;
                    item.SubItems.Add(r.Status.ToString());
                    item.SubItems.Add(r.ActualValue.ToString());
                    item.SubItems.Add(r.GoalValue.ToString());
                }
            }
        }

        /// <summary>
        /// Processes time range based SLA
        /// </summary>
        /// <param name="r">SLA time range based result</param>
        private void ProcessTransactionSLA(SlaTimeRangeRuleResult r)
        {
            slaListView.Items.Add("");
            ListViewItem item = new ListViewItem();
            if (r is SlaTransactionTimeRangeRuleResult)
                item.Text = r.Measurement.ToString() +
                    " (" + ((SlaTransactionTimeRangeRuleResult)r).TransactionName + ")";
            else
                item.Text = r.Measurement.ToString();

            item.Font = new Font(slaListView.Font, FontStyle.Bold);
            SetColorByStatus(item, r.Status);
            slaListView.Items.Add(item);      

            foreach (SlaTimeRangeInfo trResult in r.TimeRanges)
            {
                item = slaListView.Items.Add(DecodeTime(trResult.StartTime) + " -> " +
                    DecodeTime(trResult.EndTime));
                item.BackColor = Color.LightGray;
                item.SubItems.Add(trResult.Status.ToString());
                item.SubItems.Add(trResult.ActualValue.ToString());
                item.SubItems.Add(trResult.GoalValue.ToString());
            }
        }

        /// <summary>
        /// Sets ListViewItem text color accordingly to rule status
        /// </summary>
        /// <param name="item"></param>
        /// <param name="status">SLA rule status</param>
        private void SetColorByStatus(ListViewItem item, SlaRuleStatus status)
        {
            switch (status)
            {
                case SlaRuleStatus.Passed:
                    item.ForeColor = Color.Green;
                    break;
                case SlaRuleStatus.Failed:
                    item.ForeColor = Color.Red;
                    break;
                case SlaRuleStatus.NoData:
                    item.ForeColor = Color.Gray;
                    break;
            }
        }
    }

    /// <summary>
    /// Helper class used for Continuous values
    /// </summary>
    public class SimpleFilter
    {
        private string _Name;
        private double _Min;
        private double _Max;

        public SimpleFilter(string Name, double Min, double Max)
        {
            this._Min = Min;
            this._Max = Max;
            this._Name = Name;
        }

        /*
         * all properties used for binding to data grid view
         * Name - the name of continuous value
         * Min - min property of continuous value
         * Max - max property of continuous value
         **/
        public string Name
        {
            get
            {
                return _Name;
            }

            set
            {
                _Name = value;
            }
        }

        public double Min
        {
            get
            {
                return _Min;
            }

            set
            {
                _Min = value;
            }
        }

        public double Max
        {
            get
            {
                return _Max;
            }

            set
            {
                _Max = value;
            }
        }
    };
}
