using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

/// <summary>
/// Generic Function to Create CRM Service and Context
/// </summary>
namespace D365.Workflow.TFL.Helper
{
    public class CRMHelper : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);

        public static void getCRMServices(CodeActivityContext serviceProvider)
        {
            // Obtain the execution context from the service provider.  
            WorkflowContext = (IWorkflowContext)
                serviceProvider.GetExtension<IWorkflowContext>();

            // Obtain the organization service reference which you will need for  
            // web service calls.  
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetExtension<IOrganizationServiceFactory>();
            OrgService = serviceFactory.CreateOrganizationService(WorkflowContext.UserId);
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                handle.Dispose();
                // Free any other managed objects here.
                //
            }

            disposed = true;
        }

        /// <summary>
        /// Plugin Context Property
        /// </summary>
        public static IWorkflowContext WorkflowContext { get; set; }
        /// <summary>
        /// CRM Service Property
        /// </summary>
        public static IOrganizationService OrgService { get; set; }

    }
}
