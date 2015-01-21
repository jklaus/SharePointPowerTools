using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePointPowerTools.Creators.TimerJob
{
    internal class TimerJobCreator : SPJobDefinition
    {
        [Persisted]
        private string _url = "";
        [Persisted]
        private string _description = "";
        [Persisted]
        private Type _methodType = null;
        [Persisted]
        private string _methodName = null;

        internal string JobDescription
        {
            // Get accessible via 'Description' property
            set { _description = value; }
        }
        internal string Url
        {
            get { return _url; }
            set { _url = value; }
        }
        internal Type MethodType
        {
            get { return _methodType; }
            set { _methodType = value; }
        }
        internal string MethodName
        {
            get { return _methodName; }
            set { _methodName = value; }
        }

        // Overriding inheirted property - private var is set elsewhere
        public override string Description
        {
            get { return _description; }
        }

        // Using reflection to dynamically pass the executing method into the timer job instance.
        // Note that by defualt the executing method takes an SPSite argument.  If you wish to modify this, this method will need to be overridden again.
        public override void Execute(Guid contentDbId)
        {
            SPWebApplication webApplication = this.Parent as SPWebApplication;
            using (SPSite spSite = webApplication.Sites[this.Url])
            {
                if (spSite == null)
                {
                    throw new NullReferenceException(string.Format("Site ({0}) cannot be found on WebApp ({1})", this.Url, webApplication.Name));
                }

                var method = this.MethodType.GetMethod(this.MethodName);
                object classInstance = Activator.CreateInstance(this.MethodType);

                method.Invoke(classInstance, new object[] { spSite });
            }
        }

        public TimerJobCreator()
            : base()
        {
        }

        public TimerJobCreator(string Name, SPWebApplication WebApp)
            : base(Name, WebApp, null, SPJobLockType.Job)
        {
        }
    }
}
