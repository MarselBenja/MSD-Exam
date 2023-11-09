using Microsoft.Xrm.Sdk;
using System;
using System.ServiceModel;

namespace Exam
{
    public class LeadPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                Entity Target = (Entity)context.InputParameters["Target"];

                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    string Topic = Target.GetAttributeValue<string>("subject");
                    Entity newTask = new Entity("task");

                    newTask["regardingobjectid"] = new EntityReference("lead", Target.Id);
                    newTask["subject"] = "Follow Up";
                    Target["subject"] = Topic +" "+ DateTime.UtcNow.ToString("dd/MM/yyyy");//Format to Input only the Date.
                    service.Update(Target);
                  //  service.Create(newTask);

                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                    throw;
                }
            }

        }
    }
}
