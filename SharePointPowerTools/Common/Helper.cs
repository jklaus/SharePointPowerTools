using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePointPowerTools.Common
{
    class Helper
    {
        public static readonly Helper Instance = new Helper();

        private Helper()
        {

        }

        public void SetProperties(object objToSet, object objToSetFrom)
        {
            var propsToSet = objToSet.GetType().GetProperties();
            var propsToSetFrom = objToSetFrom.GetType().GetProperties();

            var props = propsToSet.Join(propsToSetFrom, pTS => pTS.Name, pTSF => pTSF.Name, (pTS, pTSF) => pTS);

            props = props.Where(p => p.Name != "LookupList" && p.Name != "LookupField");

            foreach (var prop in props)
            {
                var propVal = objToSetFrom.GetType().GetProperty(prop.Name).GetValue(objToSetFrom, null);

                if (prop.CanWrite)
                {
                    prop.SetValue(objToSet, propVal, null);
                }
            }
        }

        public UserProfile GetUserProfile(SPWeb spWeb, string AccountIdentifier)
        {
            UserProfile upUser = null;

            // Configure to get user profile
            SPServiceContext spServContext = SPServiceContext.GetContext(spWeb.Site);
            UserProfileManager upManager = new UserProfileManager(spServContext);

            // Get user profile
            if (upManager.UserExists(AccountIdentifier))
                upUser = upManager.GetUserProfile(AccountIdentifier);

            return upUser;
        }
    }
}
