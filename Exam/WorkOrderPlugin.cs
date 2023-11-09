using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.ServiceModel;

namespace Exam
{
    public class WorkorderPlugin : IPlugin
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
                    
                    ColumnSet allColumns = new ColumnSet(true);//Objekt per te marre te gjitha fushat e Entitetit.
                    EntityReference AgentER = Target.GetAttributeValue<EntityReference>("new_assignedagent");
                    Entity AgentEnt = service.Retrieve(AgentER.LogicalName, AgentER.Id, allColumns);

                    string AgentName = AgentEnt.GetAttributeValue<string>("new_agentname");
                    bool WorksOnMonday = AgentEnt.GetAttributeValue<bool>("new_isscheduledmonday");
                    bool WorksOnTuesday = AgentEnt.GetAttributeValue<bool>("new_isscheduledtuesday");
                    bool WorksOnWednesday = AgentEnt.GetAttributeValue<bool>("new_isscheduledwednesday");
                    bool WorksOnThursday = AgentEnt.GetAttributeValue<bool>("new_isscheduledthursday");
                    bool WorksOnFriday = AgentEnt.GetAttributeValue<bool>("new_isscheduledfriday");

                    int scheduleDay = Target.GetAttributeValue<OptionSetValue>("new_scheduledon").Value;

                    bool scheduledOnMatches = true;

                    switch (scheduleDay)
                    {
                        case (int)scheduledDay.Monday:
                            scheduledOnMatches = WorksOnMonday;
                            break;
                        case (int)scheduledDay.Tuesday: 
                            scheduledOnMatches = WorksOnTuesday;
                            break;
                        case (int)scheduledDay.Wednesday: 
                            scheduledOnMatches = WorksOnWednesday;
                            break;
                        case (int)scheduledDay.Thursay:
                            scheduledOnMatches = WorksOnThursday;
                            break;
                        case (int)scheduledDay.Friday:
                            scheduledOnMatches = WorksOnFriday;
                            break;
                    }

                    if (scheduledOnMatches != true)
                    {
                        throw new InvalidPluginExecutionException("Agent " + AgentName + " isn't available on that day");
                    }
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
