#region Summary
/******************************************************************************
// AUTHOR                   : Mark Nischalke 
// CREATE DATE              : 11/28/08 
// PURPOSE                  : Sync SharePoint list with database
// SPECIAL NOTES            : 
// FILE NAME                : $Workfile: Sync.cs $	
// VSS ARCHIVE              : $Archive: /WSSDatabaseSync/WSSDatabaseSync/Sync.cs $
// VERSION                  : $Revision: 2 $
// 
// EXTERNAL DEPENDENCIES    : 
// SPECIAL CHARACTERISTICS 
// OR LIMITATIONS           : 
//
// Copyright MANSoftDev © 2008 all rights reserved
// ===========================================================================
// $History: Sync.cs $
 * 
 * *****************  Version 2  *****************
 * User: Mark         Date: 11/30/08   Time: 1:24p
 * Updated in $/WSSDatabaseSync/WSSDatabaseSync
 * 
 * *****************  Version 1  *****************
 * User: Mark         Date: 11/28/08   Time: 12:56p
 * Created in $/WSSDatabaseSync/WSSDatabaseSync
//
******************************************************************************/
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using Microsoft.SharePoint.Navigation;
using System.Data;
using System.Data.SqlClient;
using System.Security.Permissions;

namespace WssDatabaseSync
{
    public sealed class Sync
    {
        /// <summary>
        /// Private ctor
        /// </summary>
        private Sync()
        {

        }

        /// <summary>
        /// Start syncing projects
        /// </summary>
        [SecurityPermission(SecurityAction.LinkDemand, Unrestricted=true)]
        public static void StartSync()
        {
            try
            {
                SyncProjects projects = new SyncProjects();
                foreach(SyncProject project in projects.Projects)
                {
                    SyncProject(project);
                }            
            }
            catch(Exception ex)
            {
                if(ExceptionPolicy.HandleException(ex, "Default Policy"))
                    throw;                
            }

        }

        #region Private Methods

        /// <summary>
        /// Sync given project
        /// </summary>
        /// <param name="project">SyncProject to process</param>
        private static void SyncProject(SyncProject project)
        {
            try
            {
                SPList list = null;
                Console.WriteLine("Attempting to open: " + project.Destination.Site);
                using(SPSite site = new SPSite(project.Destination.Site))
                {
                    using(SPWeb web = site.OpenWeb())
                    {
                        try
                        {
                            Console.WriteLine("Varify list exists...");
                            list = web.Lists[project.Destination.List];
                        }
                        catch(ArgumentException)
                        {
                            // Can't find list so just eat exception and create the list
                            Console.WriteLine("Creating list...");
                            list = CreateList(web, project.Destination.List, project.Columns);
                            CreateView(list, project.Columns);
                            CreateNavigationMenuItem(web, project.Destination.List);
                        }

                        // Should be valid by now but check anyway
                        if(list != null)
                        {
                            SyncList(project, list);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                if(ExceptionPolicy.HandleException(ex, "Default Policy"))
                    throw;
            }
        }

        /// <summary>
        /// Create list with columns
        /// </summary>
        /// <param name="web">SPWeb to add list to</param>
        /// <param name="listName">Name of list to create</param>
        /// <returns>SPList</returns>
        private static SPList CreateList(SPWeb web, string listName, List<SyncColumn> columns)
        {
            SPList list = null;
            try
            {
                Guid listGuid = web.Lists.Add(listName, listName, web.ListTemplates["Custom List"]);
                list = web.Lists[listGuid];
                SPFieldCollection fields = list.Fields;

                foreach(SyncColumn col in columns)
                {
                    SPField field = new SPField(fields, col.DataType, col.Destination);
                    list.Fields.Add(field);
                }
                
                list.Update();
                web.Update();
            }
            catch(ArgumentException ex)
            {
                Console.WriteLine("Create list failed");
                if(ExceptionPolicy.HandleException(ex, "Default Policy"))
                    throw;
            }
            return list;
        }

        /// <summary>
        /// Create view for the given list
        /// </summary>
        /// <param name="list">List to create view for</param>
        private static void CreateView(SPList list, List<SyncColumn> columns)
        {
            try
            {
                // Try to get the defualt 'All Items' view
                SPView view = list.Views["All Items"];

                // Remove the Title and Attachment columns
                if(view.ViewFields.Exists("LinkTitle"))
                    view.ViewFields.Delete("LinkTitle");

                if(view.ViewFields.Exists("Attachments"))
                    view.ViewFields.Delete("Attachments");                

                foreach(SyncColumn col in columns)
                {
                    view.ViewFields.Add(col.Destination);
                }

                view.Update();
            }
            catch(ArgumentException)
            {
                // Highly unlikely that All Items view couldn't be found
                // so just recorded it.
                Console.WriteLine("All Items view does not exist");
                Logger.Write("All Items view does not exist", "Log");
            }
        }

        /// <summary>
        /// Create navigation item fot the given list name
        /// </summary>
        /// <param name="web">SPWeb to add menu item to</param>
        /// <param name="listName">Name of list to create</param>
        private static void CreateNavigationMenuItem(SPWeb web, string listName)
        {
            // Create a navigation item for this list
            string url = "Lists/" + listName + "/AllItems.aspx";
            SPNavigationNode navNode = new SPNavigationNode(listName, url);
            foreach(SPNavigationNode node in web.Navigation.QuickLaunch)
            {
                // Find the Lists node
                if(node.Title == "Lists")
                {
                    bool menuFound = false;
                    // Check if menu item already exists
                    foreach(SPNavigationNode item in node.Children)
                    {
                        if(item.Url == navNode.Url)
                        {
                            menuFound = true;
                            break;
                        }
                    }
                    // If the menu wasn't found then add it
                    if(!menuFound)
                        node.Children.AddAsLast(navNode);
                }
            }
        }

        /// <summary>
        /// Sync the given project to the specified list
        /// </summary>
        /// <param name="project">Project to sync</param>
        /// <param name="list">List to sync with</param>
        private static void SyncList(SyncProject project, SPList list)
        {
            DataTable dt = GetDataSource(project);

            Console.Write("Creating items...");
            foreach(DataRow row in dt.Rows)
            {
                Console.Write(".");
                SPListItem item = null;

                if(project.Destination.ShouldAppend)
                    item = FindItem(list, project.Destination.DestinationKeyField, row[project.Destination.SourceKeyField]);
                else
                    item = list.Items.Add();

                foreach(SyncColumn col in project.Columns)
                {
                    if(col.Source.ToUpper().CompareTo("TODAY") == 0)
                        item[col.Destination] = GetToday(col.TimeZone);
                    else
                        item[col.Destination] = row[col.Source] == DBNull.Value ? null : row[col.Source];
                }

                item.Update();
            }
            Console.WriteLine("{0} items added", dt.Rows.Count);
        }

        /// <summary>
        /// Get datasource for project
        /// </summary>
        /// <param name="project">SyncProject to get datasource for</param>
        /// <returns>DataTable</returns>
        private static DataTable GetDataSource(SyncProject project)
        {
            Console.WriteLine("Connecting to database...");
            DataTable dt = new DataTable();
            using(SqlConnection conn = new SqlConnection(project.Source.ConnectionString))
            {
                string cmdString = string.Empty;

                if(project.Source.IsStoredProc)
                    cmdString = project.Source.DataSource;
                else
                    cmdString = "SELECT * FROM " + project.Source.DataSource;

                using(SqlCommand cmd = new SqlCommand(cmdString, conn))
                {
                    if(project.Source.IsStoredProc)
                        cmd.CommandType = CommandType.StoredProcedure;

                    conn.Open();
                    dt.Load(cmd.ExecuteReader(CommandBehavior.CloseConnection));
                }
            }
            return dt;
        }

        /// <summary>
        /// Locate an item in the list
        /// </summary>
        /// <param name="list">List to search</param>
        /// <param name="destinationField">Field to search</param>
        /// <returns>SPListItem that was found, or a new SPListItem if not found</returns>
        private static SPListItem FindItem(SPList list, string destinationField, object value)
        {
            foreach(SPListItem item in list.Items)
            {
                if(item[destinationField].ToString() == value.ToString())
                    return item;
            }

            return list.Items.Add();
        }

        /// <summary>
        /// Get current Date/Time in specified time zone
        /// </summary>
        /// <returns>String representation desired time zone</returns>
        private static string GetToday(string timeZone)
        {
            DateTime dt = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.FindSystemTimeZoneById(timeZone));
            return dt.ToString();
        }

        #endregion

        #region Properties

        #endregion
    }
}
