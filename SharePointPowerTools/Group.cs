using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SharePointPowerTools
{
    public abstract class Group : IGroup
    {
        public string Name { get; protected set; }
        public string Description { get; protected set; }
        public SPRoleType RoleType { get; protected set; }
        public string OwnerAcctName { get; protected set; }

        /// <summary>
        /// The first account name provided will serve as the default user for the group, which is added at time of creation.  If no account names are provided the SPWeb.Author will be used.
        /// If multiple account names are provided the additional names will be added to the group during the configuration process.
        /// </summary>
        public IEnumerable<string> InitialUserAcctNames { get; protected set; }

        protected Group()
        {
            RoleType = SPRoleType.Reader;
            InitialUserAcctNames = new List<String>();
        }

        /// <summary>
        /// Creates the group in the given SPWeb per the create mode settings.
        /// </summary>
        /// <param name="spWeb">Group will be added to SPWeb.SiteGroups</param>
        /// <param name="createMode">
        /// IgnoreExisting - Does nothing if group already exists.
        /// ReplaceExisting - Removes the existing group, creates and configures new group.
        /// UpdateExisting - Configures existing group per custom group settings.  Does not modify default user.
        ///  </param>
        public void Create(SPWeb spWeb, CreateMode createMode)
        {
            SPGroup spGroup = EnsureGroup(spWeb, createMode);

            // If spGroup == null, group already exists and create mode is set to ignore existing
            if (spGroup != null)
            {
                ConfigureGroup(spGroup);
            }
                        
        }

        /// <summary>
        /// Removes the group from the given SPWeb.
        /// </summary>
        /// <param name="spWeb">Group will be removed from SPWeb.SiteGroups</param>
        public void Remove(SPWeb spWeb)
        {
            var spGroup = GetAsSPGroup(spWeb);

            spWeb.Groups.RemoveByID(spGroup.ID);
            spWeb.Update();
        }

        public IEnumerable<string> GetMemberEmails(SPWeb spWeb)
        {
            var emails = new List<string>();
            
            var grp = this.GetAsSPGroup(spWeb);

            foreach(SPUser usr in grp.Users)
            {
                var up = SharePointPowerTools.Common.Helper.Instance.GetUserProfile(spWeb, usr.LoginName);

                if(up != null)
                {
                    var email = up[PropertyConstants.WorkEmail].Value.ToString();
                    emails.Add(email);
                }
            }

            return emails;
        }

        public void EnsureSPUserInGroup(SPWeb spWeb, SPUser spUser)
        {
            var spGroup = this.GetAsSPGroup(spWeb);

            if (spGroup.Users.Cast<SPUser>().Where(u => u.LoginName == spUser.LoginName).Count() < 1)
            {
                spGroup.AddUser(spUser);
                spGroup.Update();
            }
        }

        private void ConfigureGroup(SPGroup spGroup)
        {
            SPWeb spWeb = spGroup.ParentWeb;

            //SPRoleDefinition oRole = spWeb.RoleDefinitions.GetByType(this.RoleType);
            SPRoleDefinition oRole = spWeb.RoleDefinitions.Cast<SPRoleDefinition>().Where(rd => rd.Type.ToString() == this.RoleType.ToString()).Last();
            SPRoleAssignment oRoleAssignment = new SPRoleAssignment(spGroup);
            oRoleAssignment.RoleDefinitionBindings.Add(oRole);
            spWeb.RoleAssignments.Add(oRoleAssignment);

            spGroup.Description = this.Description;

            SPMember spOwner = spWeb.EnsureUser(string.IsNullOrEmpty(this.OwnerAcctName) ? spWeb.Author.LoginName : this.OwnerAcctName);
            spGroup.Owner = spOwner;

            AddUsersToGroup(spGroup);

            spGroup.Update();
            spWeb.Update();
        }

        private void AddUsersToGroup(SPGroup spGroup)
        {
            var spWeb = spGroup.ParentWeb;
            foreach (var userName in this.InitialUserAcctNames)
            {
                var user = spWeb.EnsureUser(userName);
                // Ensure user does not already exist in group, add if necessary.
                if (!spGroup.Users.Cast<SPUser>().Select(u=>u.ID).Contains(user.ID))
                {
                    spGroup.AddUser(user);
                }
            }
        }

        private SPGroup EnsureGroup(SPWeb spWeb, CreateMode createMode)
        {
            SPGroup spGroup = GetAsSPGroup(spWeb);

            if (createMode == CreateMode.ReplaceExisting && spGroup != null)
            {
                Remove(spWeb);
                spGroup = AddGroup(spWeb);
            }
            else if (createMode == CreateMode.IgnoreExisting && spGroup != null)
            {
                return null;
            }
            else if (spGroup == null)
            {
                spGroup = AddGroup(spWeb);
            }

            return spGroup;
        }

        private SPGroup AddGroup(SPWeb spWeb)
        {
            bool webSafety = spWeb.AllowUnsafeUpdates;
            spWeb.AllowUnsafeUpdates = true;

            SPMember spOwner = spWeb.EnsureUser(string.IsNullOrEmpty(this.OwnerAcctName) ? spWeb.Author.LoginName : this.OwnerAcctName);
            SPUser spDefUser = spWeb.EnsureUser(string.IsNullOrEmpty(this.InitialUserAcctNames.FirstOrDefault()) ? spWeb.Author.LoginName : this.InitialUserAcctNames.First());

            // Disabled for testing purposes - Not sure if this will be used, conditions statement in Ensure should prevent this from ever being an issue
            // may stay with throughing exception rather than delete existing.
            //SPGroup spGroup = GetAsSPGroup(spWeb);

            //if (spGroup != null)
            //{
            //    spWeb.SiteGroups.RemoveByID(spGroup.ID);
            //}

            spWeb.SiteGroups.Add(this.Name, spOwner, spDefUser, this.Description);

            spWeb.AllowUnsafeUpdates = webSafety;
            spWeb.Update();

            SPGroup spGroup = GetAsSPGroup(spWeb);
            return spGroup;
        }
        
        /// <summary>
        /// Returns the SPGroup associated with the custom group in the SPWeb context.
        /// Returns null if the group does not exist.
        /// </summary>
        /// <param name="spWeb">Used to access SPWeb.SiteGroups.</param>
        /// <returns>Returns the SPGroup if it exists, if not returns null.</returns>
        public SPGroup GetAsSPGroup(SPWeb spWeb)
        {
            SPGroup spGroup = spWeb.SiteGroups.Cast<SPGroup>().Where(g => g.Name == this.Name).FirstOrDefault();
            return spGroup;
        }

        /// <summary>
        /// Returns an IEnumerable of string Corp IDs associated with the group members.
        /// </summary>
        /// <param name="spWeb"></param>
        /// <returns>Returns an IEnumerable of string Corp IDs</returns>
        public IEnumerable<string> GetMemeberCorpIds(SPWeb spWeb)
        {
            var spGroup = this.GetAsSPGroup(spWeb);

            var users = spGroup.Users.Cast<SPUser>();

            var userIds = users.Select(u => Regex.Replace(u.LoginName, ".*[\\\\]", ""));

            return userIds;
        }
    }
}
