using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using SharePointPowerTools.Creators.TimerJob;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePointPowerTools
{
    public abstract class TimerJob
    {
        private string _title = null;
        private string _name = null;
        
        public string Description { get; protected set; }
        public string Name //{ get; protected set; }
        {
            get
            {
                return _name;
            }
            protected set
            {
                _name = value.Replace(" ", "_");
            }
        }

        public string Title // get; protected set;
        {
            get
            {
                if (string.IsNullOrEmpty(_title))
                {
                    return Name;
                }
                else
                {
                    return _title;
                }
            }
            protected set
            {
                _title = value;
            }
        }
        public Action<SPSite> ExecutingMethod { get; protected set; } // 
        public string Url { get; protected set; }
        public SPSchedule Schedule { get; protected set; }
       

        public TimerJob()
        {
        }

        public void Create(SPWebApplication spWebApp)
        {
            //Ensure timer job will only be added to webapp associated with specified site.
            if (spWebApp.Sites[this.Url] != null)
            {
                ValidateTitleAndName();

                RemoveExistingJobs(spWebApp);

                var job = new TimerJobCreator(this.Name, spWebApp)
                {
                    Title = this.Title,
                    Name = this.Name,
                    JobDescription = this.Description,
                    Schedule = this.Schedule,
                    Url = this.Url,
                    MethodName = this.ExecutingMethod.Method.Name,
                    MethodType = this.ExecutingMethod.Method.DeclaringType
                };

                job.Update();
            }
            else
            {
                throw new NullReferenceException(string.Format("Site ({0}) cannot be found on WebApp ({1})", this.Url, spWebApp.Name));
            }
        }

        public void Remove(SPWebApplication spWebApp)
        {
            // Ensure timer job will only be removed from webapp associated with specified site.
            if (spWebApp.Sites[this.Url] != null)
            {
                ValidateTitleAndName();
                RemoveExistingJobs(spWebApp);
            }
        }

        private void RemoveExistingJobs(SPWebApplication spWebApp)
        {
            foreach (SPJobDefinition job in spWebApp.JobDefinitions)
            {
                if (job.Name == this.Name)
                {
                    job.Delete();
                }
            }
        }

        private void ValidateTitleAndName()
        {
            if (!this.Title.Contains(this.Url))
            {
                this.Title = string.Format("{0} - {1}", this.Title, this.Url);
            }

            if (!this.Name.Contains(this.Url))
            {
                this.Name = string.Format("{0}_{1}", this.Name, this.Url);
            }
        }
    }
}
