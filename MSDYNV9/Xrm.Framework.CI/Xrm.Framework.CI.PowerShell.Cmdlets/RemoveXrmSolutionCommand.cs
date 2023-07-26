using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Xrm.Framework.CI.Common;
using Xrm.Framework.CI.Common.Entities;
using Xrm.Framework.CI.PowerShell.Cmdlets.Common;

namespace Xrm.Framework.CI.PowerShell.Cmdlets
{
    /// <summary>
    /// <para type="synopsis">Removes a CRM Solution.</para>
    /// <para type="description">The Remove-XrmSolution of a CRM solution by unique name.
    /// </para>
    /// </summary>
    /// <example>
    ///   <code>C:\PS>Remove-XrmSolution -ConnectionString "" -UniqueSolutionName "UniqueSolutionName"</code>
    ///   <para>Exports the "" managed solution to "" location</para>
    /// </example>
    [Cmdlet(VerbsCommon.Remove, "XrmSolution")]
    [OutputType(typeof(String))]
    public class RemoveXrmSolutionCommand : XrmCommandBase
    {
        #region Parameters

        /// <summary>
        /// <para type="description">The unique name of the solution components to be removed.</para>
        /// </summary>
        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        #endregion

        #region Process Record

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            base.WriteVerbose(string.Format("Removing Solution: {0}", SolutionName));

            using (var context = new CIContext(OrganizationService))
            {
                var query1 = from solution in context.SolutionSet
                             where solution.UniqueName == SolutionName
                             select solution.Id;

                if (query1 == null)
                {
                    throw new Exception(string.Format("Solution {0} could not be found", SolutionName));
                }

                XrmConnectionManager xrmConnection = new XrmConnectionManager(Logger);
                IOrganizationService pollingOrganizationService = xrmConnection.Connect(ConnectionString, 120);
                SolutionManager solutionManager = new SolutionManager(Logger, OrganizationService, pollingOrganizationService);

                solutionManager.DeleteSolution(SolutionName);
            }

            WriteVerbose($"Removed Solution {SolutionName}");
        }

        #endregion
    }
}