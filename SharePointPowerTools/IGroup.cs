using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePointPowerTools
{   
    public interface IGroup
    {
        string Name { get; }
        string Description { get; }
        SPRoleType RoleType { get; }

        /// <summary>
        /// The first account name provided will serve as the default user for the group, which is added at time of creation.  If no account names are provided the SPWeb.Author will be used.
        /// If multiple account names are provided the additional names will be added to the group during the configuration process.
        /// </summary>
        IEnumerable<string> InitialUserAcctNames { get; }
        string OwnerAcctName { get; }

        void Create(SPWeb spWeb, CreateMode createMode);
        void Remove(SPWeb spWeb);
    }
}
