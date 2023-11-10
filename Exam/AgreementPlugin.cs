using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.ServiceModel;

namespace Exam
{
    public class AgreementPlugin : IPlugin
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
                    EntityReference AccountER = Target.GetAttributeValue<EntityReference>("cr260_account");
                    Entity Account = service.Retrieve(AccountER.LogicalName, AccountER.Id, new ColumnSet("primarycontactid"));
                    EntityReference Contact = Account.GetAttributeValue<EntityReference>("primarycontactid");
                    int AgreementType = Target.GetAttributeValue<OptionSetValue>("cr260_agreementtype").Value;

                    switch (AgreementType)
                    {
                        case (int)agreementType.Onboarding:
                            //Check if any Onboarding Agreements exists with selected Account.
                            var fetchOnboarding = $@"<fetch version='1.0' mapping='logical' no-lock='false' distinct='true'>
                                 <entity name='cr260_agreement'>
                                 <attribute name='cr260_account'/>
                                 <attribute name='cr260_agreementtype'/>
                                  <filter type='and'>
                                  <condition attribute='statecode' operator='eq' value='{(int)stateCode.Active}'/>
                                  <condition attribute='cr260_account' operator='eq' value='{AccountER.Id}'/>
                                  <condition attribute='cr260_agreementtype' operator='eq' value='{(int)agreementType.Onboarding}'/>
                                </filter>
                                </entity>
                                </fetch>";
                            EntityCollection Onboardings = service.RetrieveMultiple(new FetchExpression(fetchOnboarding));
     
                            if (Onboardings.Entities.Count > 0)
                            {
                                throw new InvalidPluginExecutionException("There already is an agreement of type Onboarding associated with this Account");
                            }
                            else if (Target.Contains("cr260_agreementstartdate") && Target.Contains("cr260_agreementenddate"))
                            {
                                //Create Query with filter Account Id and Update Field "cr260_tc" on Opportunities.
                                QueryExpression query = new QueryExpression("opportunity");
                                query.ColumnSet = new ColumnSet("parentaccountid", "statecode");

                                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                                ConditionExpression stateCondition = new ConditionExpression("statecode", ConditionOperator.Equal, (int)stateCode.Active);
                                ConditionExpression accountCondition = new ConditionExpression("parentaccountid", ConditionOperator.Equal, Account.Id);
                                filter.AddCondition(stateCondition);
                                filter.AddCondition(accountCondition);

                                query.Criteria = filter;

                                EntityCollection Opportunities = service.RetrieveMultiple(query);

                                foreach(var Opportunity in Opportunities.Entities){

                                    Entity OpportunityEnt = new Entity(Opportunity.LogicalName, Opportunity.Id);
                                    bool tc = Opportunity.GetAttributeValue<bool>("cr260_tc");

                                    Opportunity["cr260_tc"] = true;
                                    service.Update(Opportunity);
                                    Console.WriteLine("Updated Records" + Opportunity.Id.ToString());
                                }
                            }
                            break;
                        case (int)agreementType.NDA:
                            //Check if any NDAs Agreements exists with selected Account.
                            var fetchNDA = $@"<fetch version='1.0' mapping='logical' no-lock='false' distinct='true'>
                                 <entity name='cr260_agreement'>
                                 <attribute name='cr260_account'/>
                                  <attribute name='cr260_agreementtype'/>
                                      <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='{(int)stateCode.Active}'/>
                                      <condition attribute='cr260_account' operator='eq' value='{AccountER.Id}'/>
                                      <condition attribute='cr260_agreementtype' operator='eq' value='{(int)agreementType.NDA}'/>
                                   </filter>
                                  </entity>
                                </fetch>";
                            EntityCollection NDA = service.RetrieveMultiple(new FetchExpression(fetchNDA));
                           
                            if (NDA.Entities.Count > 0)
                            {
                                throw new InvalidPluginExecutionException("There already is an agreement of type NDA associated with this Account");
                            }
                            break;
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
