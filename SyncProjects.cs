#region Summary
/******************************************************************************
// AUTHOR                   : Mark Nischalke 
// CREATE DATE              : 11/28/08 
// PURPOSE                  : Classes to represent data sync elements
// SPECIAL NOTES            : 
// FILE NAME                : $Workfile: SyncProjects.cs $	
// VSS ARCHIVE              : $Archive: /WSSDatabaseSync/WSSDatabaseSync/SyncProjects.cs $
// VERSION                  : $Revision: 2 $
// 
// EXTERNAL DEPENDENCIES    : 
// SPECIAL CHARACTERISTICS 
// OR LIMITATIONS           : 
//
// Copyright MANSoftDev © 2008 all rights reserved
// ===========================================================================
// $History: SyncProjects.cs $
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
using System.Configuration;
using System.Xml.Linq;
using System.IO;

namespace WssDatabaseSync
{
    public class SyncProjects
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public SyncProjects()
        {
            if(File.Exists(ProjectsFile))
                ProjectsDocument = XElement.Load(ProjectsFile);
            else
                throw new FileNotFoundException(ProjectsFile);

            LoadProjects();
        }

        #region Private Methods

        /// <summary>
        /// Create SyncProjects from XML
        /// </summary>
        private void LoadProjects()
        {
            Projects = new List<SyncProject>();

            foreach(XElement projectElement in ProjectsDocument.Elements("syncProject"))
            {
                Projects.Add(new SyncProject(projectElement));
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Get path to file containing definitions
        /// of sync projects
        /// </summary>
        private static string ProjectsFile
        {
            get
            {
                if(ConfigurationManager.AppSettings["DataSyncFile"] != null)
                    return ConfigurationManager.AppSettings["DataSyncFile"];
                else
                    return "DataSync.xml";
            }
        }

        /// <summary>
        /// Get/Set XElement containing project definition
        /// </summary>
        private XElement ProjectsDocument { get; set; }

        /// <summary>
        /// Get list of SyncProjects
        /// </summary>
        internal List<SyncProject> Projects { get; private set; }

        #endregion
    }

    #region Internal classes

    internal class SyncProject
    {
        public SyncProject(XElement projectElement)
        {
            Source = new SyncSource(projectElement.Element("syncSource"));
            Destination = new SyncDestination(projectElement.Element("syncDestination"));
            Columns = new List<SyncColumn>();
            foreach(XElement col in projectElement.Element("columns").Elements("column"))
            {
                Columns.Add(new SyncColumn(col));
            }
        }

        public SyncSource Source { get; private set; }
        public SyncDestination Destination { get; private set; }
        public List<SyncColumn> Columns { get; private set; }
    }

    internal class SyncSource
    {
        public SyncSource(XElement sourceElement)
        {
            ConnectionString = sourceElement.Element("connectionString").Value;
            DataSource = sourceElement.Element("source").Value;
            IsStoredProc = Convert.ToBoolean(sourceElement.Element("source").Attribute("isStoredProc").Value);
        }

        public string ConnectionString { get; private set; }
        public string DataSource { get; private set; }
        public bool IsStoredProc { get; private set; }
    }

    internal class SyncDestination
    {
        public SyncDestination(XElement sourceElement)
        {
            Site = sourceElement.Element("site").Value;
            List = sourceElement.Element("list").Value;
            ShouldAppend = Convert.ToBoolean(sourceElement.Attribute("append").Value);
            if(ShouldAppend)
            {
                SourceKeyField = sourceElement.Attribute("sourceKeyField").Value;
                DestinationKeyField = sourceElement.Attribute("destinationKeyField").Value;
            }
        }

        public string Site { get; private set; }
        public string List { get; private set; }
        public bool ShouldAppend { get; private set; }
        public string SourceKeyField { get; private set; }
        public string DestinationKeyField { get; private set; }
    }

    internal class SyncColumn
    {
        public SyncColumn(XElement sourceElement)
        {
            Source = sourceElement.Attribute("source").Value;
            Destination = sourceElement.Attribute("destination").Value;
            DataType = sourceElement.Attribute("dataType").Value;
            if(sourceElement.Attribute("timeZone") != null)
                TimeZone = sourceElement.Attribute("timeZone").Value;
            else
                TimeZone = "Eastern Standard Time";
        }

        public string Source { get; private set; }
        public string Destination { get; private set; }
        public string DataType { get; private set; }
        public string TimeZone { get; private set; }
    }

    #endregion
}