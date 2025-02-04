using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Taskbridge.Nixon.Industrial.Opportunity
{
    public class CalculateRollup : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = null;
            IOrganizationServiceFactory factory = null;
            IOrganizationService service = null;
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);

                if (service == null) return;
                if (context.InputParameters.Contains("Target"))
                {

                    if ((context.InputParameters["Target"] is Entity && Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_plannedmaintenanceservice")) ||
                        (context.InputParameters["Target"] is EntityReference && Equals(((EntityReference)context.InputParameters["Target"]).LogicalName, "bolt_plannedmaintenanceservice")))
                    {
                        Entity entPMS = null;
                        string deleteFilter = null;

                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            entPMS = (Entity)context.InputParameters["Target"];
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {
                            entPMS = (Entity)context.PostEntityImages["Image"];

                        }
                        else
                        {
                            entPMS = (Entity)context.PreEntityImages["Image"];
                            deleteFilter = String.Format("<condition attribute = 'bolt_plannedmaintenanceserviceid' operator= 'ne' uitype = 'bolt_plannedmaintenanceservice' value = '{0}' />", entPMS.Id);
                        }

                        EntityReference erfOpportunity = entPMS.Contains("bolt_serviceid") ? entPMS.GetAttributeValue<EntityReference>("bolt_serviceid") : null;
                        string groupby1 = String.Format(@"
                                        <fetch distinct='false' mapping='logical' aggregate='true'> 
                                           <entity name='bolt_plannedmaintenanceservice'>  
                                            <attribute name='bolt_totalcontractamount' alias='totalcontractamount' aggregate='sum' />
                                            <filter>
                                              <condition attribute='bolt_serviceid' operator='eq' value='{0}' />
                                               {1}
                                            </filter>
                                           </entity> 
                                        </fetch>", erfOpportunity.Id, deleteFilter);

                        EntityCollection groupby1_result = service.RetrieveMultiple(new FetchExpression(groupby1));

                        foreach (var c in groupby1_result.Entities)
                        {
                            decimal totalcontractamount = (Money)((AliasedValue)c["totalcontractamount"]).Value != null ? ((Money)((AliasedValue)c["totalcontractamount"]).Value).Value : 0;
                            Entity objOpty = new Entity(erfOpportunity.LogicalName);
                            objOpty.Id = erfOpportunity.Id;
                            objOpty["bolt_annualcontractprice_new"] = new Money(totalcontractamount);
                            service.Update(objOpty);

                        }
                    }

                    if ((context.InputParameters["Target"] is Entity && Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_kdservicemaintenance")) ||
                        (context.InputParameters["Target"] is EntityReference && Equals(((EntityReference)context.InputParameters["Target"]).LogicalName, "bolt_kdservicemaintenance")))
                    //if (Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_kdservicemaintenance"))
                    {
                        Entity entKDS = null;
                        string deleteFilter = null;

                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            entKDS = (Entity)context.InputParameters["Target"];
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {
                            entKDS = (Entity)context.PostEntityImages["Image"];
                        }
                        else
                        {
                            entKDS = (Entity)context.PreEntityImages["Image"];
                            deleteFilter = String.Format("<condition attribute = 'bolt_kdservicemaintenanceid' operator= 'ne' uitype = 'bolt_kdservicemaintenance' value = '{0}' />", entKDS.Id);
                        }

                        EntityReference erfOpportunity = entKDS.Contains("bolt_kdserviceid") ? entKDS.GetAttributeValue<EntityReference>("bolt_kdserviceid") : null;
                        string groupby1 = String.Format(@"
                                        <fetch distinct='false' mapping='logical' aggregate='true'> 
                                           <entity name='bolt_kdservicemaintenance'>  
                                            <attribute name='bolt_totalkdcontractprice' alias='totalkdcontractprice' aggregate='sum' />
                                            <filter>
                                              <condition attribute='bolt_kdserviceid' operator='eq' value='{0}' />  
                                              {1}
                                            </filter>
                                           </entity> 
                                        </fetch>", erfOpportunity.Id, deleteFilter);

                        EntityCollection groupby1_result = service.RetrieveMultiple(new FetchExpression(groupby1));

                        foreach (var c in groupby1_result.Entities)
                        {
                            decimal totalkdcontractprice = (Money)((AliasedValue)c["totalkdcontractprice"]).Value != null ? ((Money)((AliasedValue)c["totalkdcontractprice"]).Value).Value : 0;
                            Entity objOpty = new Entity(erfOpportunity.LogicalName);
                            objOpty.Id = erfOpportunity.Id;
                            objOpty["bolt_kdcontractprice_new"] = new Money(totalkdcontractprice);
                            service.Update(objOpty);

                        }
                    }

                    if ((context.InputParameters["Target"] is Entity && Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_miscitems")) ||
                        (context.InputParameters["Target"] is EntityReference && Equals(((EntityReference)context.InputParameters["Target"]).LogicalName, "bolt_miscitems")))
                    //if (Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_miscitems"))
                    {
                        Entity entMisc = null;
                        string deleteFilter = null;

                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            entMisc = (Entity)context.InputParameters["Target"];
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {
                            entMisc = (Entity)context.PostEntityImages["Image"];
                        }
                        else
                        {
                            entMisc = (Entity)context.PreEntityImages["Image"];
                            deleteFilter = String.Format("<condition attribute = 'bolt_miscitemsid' operator= 'ne' uitype = 'bolt_miscitems' value = '{0}' />", entMisc.Id);
                        }
                        EntityReference erfOpportunity = entMisc.Contains("bolt_miscitemid") ? entMisc.GetAttributeValue<EntityReference>("bolt_miscitemid") : null;
                        string groupby1 = String.Format(@"
                                        <fetch distinct='false' mapping='logical' aggregate='true'> 
                                           <entity name='bolt_miscitems'>  
                                            <attribute name='bolt_miscprice' alias='miscprice' aggregate='sum' />
                                            <filter>
                                              <condition attribute='bolt_miscitemid' operator='eq' value='{0}' />
                                              {1}
                                            </filter>
                                           </entity> 
                                        </fetch>", erfOpportunity.Id, deleteFilter);

                        EntityCollection groupby1_result = service.RetrieveMultiple(new FetchExpression(groupby1));

                        foreach (var c in groupby1_result.Entities)
                        {
                            decimal miscprice = (Money)((AliasedValue)c["miscprice"]).Value != null ? ((Money)((AliasedValue)c["miscprice"]).Value).Value : 0;
                            Entity objOpty = new Entity(erfOpportunity.LogicalName);
                            objOpty.Id = erfOpportunity.Id;
                            objOpty["bolt_miscitemcontractprice_new"] = new Money(miscprice);
                            service.Update(objOpty);
                        }
                    }

                    if ((context.InputParameters["Target"] is Entity && Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_specialpricingservicemaintenance")) ||
                        (context.InputParameters["Target"] is EntityReference && Equals(((EntityReference)context.InputParameters["Target"]).LogicalName, "bolt_specialpricingservicemaintenance")))
                    //if (Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_specialpricingservicemaintenance"))
                    {
                        Entity entSPSI = null;
                        string deleteFilter = null;

                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            entSPSI = (Entity)context.InputParameters["Target"];
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {
                            entSPSI = (Entity)context.PostEntityImages["Image"];
                        }
                        else
                        {
                            entSPSI = (Entity)context.PreEntityImages["Image"];
                            deleteFilter = String.Format("<condition attribute = 'bolt_specialpricingservicemaintenanceid' operator= 'ne' uitype = 'bolt_specialpricingservicemaintenance' value = '{0}' />", entSPSI.Id);
                        }
                        EntityReference erfOpportunity = entSPSI.Contains("bolt_specialservicepricingid") ? entSPSI.GetAttributeValue<EntityReference>("bolt_specialservicepricingid") : null;
                        string groupby1 = String.Format(@"
                                        <fetch distinct='false' mapping='logical' aggregate='true'> 
                                           <entity name='bolt_specialpricingservicemaintenance'>  
                                            <attribute name='bolt_totalcontractamount' alias='totalcontractamount' aggregate='sum' />
                                            <filter>
                                              <condition attribute='bolt_specialservicepricingid' operator='eq' value='{0}' />
                                              {1}
                                            </filter>
                                           </entity> 
                                        </fetch>", erfOpportunity.Id, deleteFilter);

                        EntityCollection groupby1_result = service.RetrieveMultiple(new FetchExpression(groupby1));

                        foreach (var c in groupby1_result.Entities)
                        {
                            decimal totalcontractamount = (Money)((AliasedValue)c["totalcontractamount"]).Value != null ? ((Money)((AliasedValue)c["totalcontractamount"]).Value).Value : 0;
                            Entity objOpty = new Entity(erfOpportunity.LogicalName);
                            objOpty.Id = erfOpportunity.Id;
                            objOpty["bolt_specialpricingcontractamount_new"] = new Money(totalcontractamount);
                            service.Update(objOpty);

                        }
                    }

                    if ((context.InputParameters["Target"] is Entity && Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_atsplannedmaintenance")) ||
                        (context.InputParameters["Target"] is EntityReference && Equals(((EntityReference)context.InputParameters["Target"]).LogicalName, "bolt_atsplannedmaintenance")))
                    {
                        Entity entATSPM = null;
                        string deleteFilter = null;

                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            entATSPM = (Entity)context.InputParameters["Target"];
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {
                            entATSPM = (Entity)context.PostEntityImages["Image"];
                        }
                        else
                        {
                            entATSPM = (Entity)context.PreEntityImages["Image"];
                            deleteFilter = String.Format("<condition attribute = 'bolt_atsplannedmaintenanceid' operator= 'ne' uitype = 'bolt_atsplannedmaintenance' value = '{0}' />", entATSPM.Id);
                        }

                        EntityReference erfOpportunity = entATSPM.Contains("bolt_opportunityid") ? entATSPM.GetAttributeValue<EntityReference>("bolt_opportunityid") : null;
                        string groupby1 = String.Format(@"
                                        <fetch distinct='false' mapping='logical' aggregate='true'> 
                                           <entity name='bolt_atsplannedmaintenance'>  
                                            <attribute name='bolt_serviceprice' alias='serviceprice' aggregate='sum' />
                                            <filter>
                                              <condition attribute='bolt_opportunityid' operator='eq' value='{0}' />
                                               {1}
                                            </filter>
                                           </entity> 
                                        </fetch>", erfOpportunity.Id, deleteFilter);

                        EntityCollection groupby1_result = service.RetrieveMultiple(new FetchExpression(groupby1));

                        foreach (var c in groupby1_result.Entities)
                        {
                            decimal serviceprice = (Money)((AliasedValue)c["serviceprice"]).Value != null ? ((Money)((AliasedValue)c["serviceprice"]).Value).Value : 0;
                            Entity objOpty = new Entity(erfOpportunity.LogicalName);
                            objOpty.Id = erfOpportunity.Id;
                            objOpty["bolt_totalatspmamount"] = new Money(serviceprice);
                            service.Update(objOpty);

                        }
                    }


                    if ((context.InputParameters["Target"] is Entity && Equals(((Entity)context.InputParameters["Target"]).LogicalName, "bolt_remotemonitoring")) ||
                        (context.InputParameters["Target"] is EntityReference && Equals(((EntityReference)context.InputParameters["Target"]).LogicalName, "bolt_remotemonitoring")))
                    {
                        Entity entRM = null;
                        string deleteFilter = null;

                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            entRM = (Entity)context.InputParameters["Target"];
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {
                            entRM = (Entity)context.PostEntityImages["Image"];
                        }
                        else
                        {
                            entRM = (Entity)context.PreEntityImages["Image"];
                            deleteFilter = String.Format("<condition attribute = 'bolt_remotemonitoringid' operator= 'ne' uitype = 'bolt_remotemonitoring' value = '{0}' />", entRM.Id);
                        }

                        EntityReference erfOpportunity = entRM.Contains("bolt_opportunity") ? entRM.GetAttributeValue<EntityReference>("bolt_opportunity") : null;
                        string groupby1 = String.Format(@"
                                        <fetch distinct='false' mapping='logical' aggregate='true'> 
                                           <entity name='bolt_remotemonitoring'>  
                                            <attribute name='bolt_remotemonitoringserviceprice' alias='remotemonitoringserviceprice' aggregate='sum' />
                                            <filter>
                                              <condition attribute='bolt_opportunity' operator='eq' value='{0}' />
                                               {1}
                                            </filter>
                                           </entity> 
                                        </fetch>", erfOpportunity.Id, deleteFilter);

                        EntityCollection groupby1_result = service.RetrieveMultiple(new FetchExpression(groupby1));

                        foreach (var c in groupby1_result.Entities)
                        {
                            decimal serviceprice = (Money)((AliasedValue)c["remotemonitoringserviceprice"]).Value != null ? ((Money)((AliasedValue)c["remotemonitoringserviceprice"]).Value).Value : 0;
                            Entity objOpty = new Entity(erfOpportunity.LogicalName);
                            objOpty.Id = erfOpportunity.Id;
                            objOpty["new_remotemonitoringserviceprice"] = new Money(serviceprice);
                            service.Update(objOpty);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
