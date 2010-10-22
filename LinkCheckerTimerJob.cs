using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace MarabouStork.Sharepoint.LinkChecker
{
    public class LinkCheckerTimerJob : SPJobDefinition
    {
        public LinkCheckerTimerJob() : base() { }

        public LinkCheckerTimerJob(string jobName, SPWebApplication webapp) : base(jobName, webapp, null, SPJobLockType.Job) { }

        public override void Execute(Guid targetInstanceId)
        {
            base.Execute(targetInstanceId);

            foreach (SPSite siteCollection in this.WebApplication.Sites)
            {
                LinkChecker.ValidateDocumentUrls(siteCollection.Url, this.WebApplication, DateTime.MinValue);
            }
        }
    }
}

