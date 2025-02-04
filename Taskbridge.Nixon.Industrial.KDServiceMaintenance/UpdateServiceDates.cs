using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Taskbridge.Nixon.Industrial.KDSMandPMService
{
    public class UpdateServiceDates : IPlugin
    {
        IOrganizationService service;
        ITracingService tracingService;

        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parmameters.
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.LogicalName == "bolt_plannedmaintenanceservice" || entity.LogicalName == "bolt_kdservicemaintenance")
                {
                    try
                    {
                        if (context.PostEntityImages.Contains("Image"))
                        {
                            Entity objPMS = null;
                            Entity PMSDObj = null;
                            DateTime serviceStartDate = DateTime.MinValue;
                            Entity postImagePMS = (Entity)context.PostEntityImages["Image"];
                            if (entity.LogicalName == "bolt_plannedmaintenanceservice")
                            {
                                if (!postImagePMS.Attributes.Contains("bolt_pmservicedescription") || !postImagePMS.Attributes.Contains("bolt_servicestartdate"))
                                    return;
                                PMSDObj = service.Retrieve("bolt_pmservicedescription", postImagePMS.GetAttributeValue<EntityReference>("bolt_pmservicedescription").Id, new ColumnSet("bolt_name"));
                                serviceStartDate = postImagePMS.GetAttributeValue<DateTime>("bolt_servicestartdate");
                                objPMS = new Entity(entity.LogicalName, postImagePMS.Id);
                            }
                            else if (entity.LogicalName == "bolt_kdservicemaintenance")
                            {
                                if (!postImagePMS.Attributes.Contains("bolt_servicedescription") || !postImagePMS.Attributes.Contains("bolt_kdservicestartdate"))
                                    return;
                                PMSDObj = service.Retrieve("bolt_pmservicedescription", postImagePMS.GetAttributeValue<EntityReference>("bolt_servicedescription").Id, new ColumnSet("bolt_name"));
                                serviceStartDate = postImagePMS.GetAttributeValue<DateTime>("bolt_kdservicestartdate");
                                objPMS = new Entity(entity.LogicalName, postImagePMS.Id);
                            }

                            string PMSName = PMSDObj.Contains("bolt_name") ? PMSDObj.GetAttributeValue<string>("bolt_name") : "";
                            string service1Date = serviceStartDate.ToString("yyyy-MM");
                            string service2Date = serviceStartDate.AddMonths(1).ToString("yyyy-MM");
                            string service3Date = serviceStartDate.AddMonths(2).ToString("yyyy-MM");
                            string service4Date = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                            string service5Date = serviceStartDate.AddMonths(4).ToString("yyyy-MM");
                            string service6Date = serviceStartDate.AddMonths(5).ToString("yyyy-MM");
                            string service7Date = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                            string service8Date = serviceStartDate.AddMonths(7).ToString("yyyy-MM");
                            string service9Date = serviceStartDate.AddMonths(8).ToString("yyyy-MM");
                            string service10Date = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                            string service11Date = serviceStartDate.AddMonths(10).ToString("yyyy-MM");
                            string service12Date = serviceStartDate.AddMonths(11).ToString("yyyy-MM");

                            objPMS["bolt_service1type"] = null;
                            objPMS["bolt_service2type"] = null;
                            objPMS["bolt_service3type"] = null;
                            objPMS["bolt_service4type"] = null;
                            objPMS["bolt_service5type"] = null;
                            objPMS["bolt_service6type"] = null;
                            objPMS["bolt_service7type"] = null;
                            objPMS["bolt_service8type"] = null;
                            objPMS["bolt_service9type"] = null;
                            objPMS["bolt_service10type"] = null;
                            objPMS["bolt_service11type"] = null;
                            objPMS["bolt_service12type"] = null;
                            objPMS["bolt_service1datetext"] = null;
                            objPMS["bolt_service2datetext"] = null;
                            objPMS["bolt_service3datetext"] = null;
                            objPMS["bolt_service4datetext"] = null;
                            objPMS["bolt_service5datetext"] = null;
                            objPMS["bolt_service6datetext"] = null;
                            objPMS["bolt_service7datetext"] = null;
                            objPMS["bolt_service8datetext"] = null;
                            objPMS["bolt_service9datetext"] = null;
                            objPMS["bolt_service10datetext"] = null;
                            objPMS["bolt_service11datetext"] = null;
                            objPMS["bolt_service12datetext"] = null;
                            objPMS["bolt_service1date"] = null;
                            objPMS["bolt_service2date"] = null;
                            objPMS["bolt_service3date"] = null;
                            objPMS["bolt_service4date"] = null;
                            objPMS["bolt_service5date"] = null;
                            objPMS["bolt_service6date"] = null;
                            objPMS["bolt_service7date"] = null;
                            objPMS["bolt_service8date"] = null;
                            objPMS["bolt_service9date"] = null;
                            objPMS["bolt_service10date"] = null;
                            objPMS["bolt_service11date"] = null;
                            objPMS["bolt_service12date"] = null;

                            if (PMSName.Equals("ANNUAL (MAJOR)") || PMSName.Equals("GREEN ANNUAL") || PMSName.Equals("GREEN ANNUAL (W GREEN SVC)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                            }

                            else if (PMSName.Equals("SEMI-ANNUAL GREEN") || PMSName.Equals("SEMI-ANNUAL GREEN (W GREEN SVC)") || PMSName.Equals("SEMI-ANNUAL MAJOR (W MAJ SVC)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(6);
                            }

                            else if (PMSName.Equals("SEMI-ANNUAL MINOR (W MAJ SVC)") || PMSName.Equals("SEMI-ANNUAL MINOR (W GREEN SVC)") || PMSName.Equals("SEMI-ANNUAL MINOR"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(6);
                            }

                            else if (PMSName.Equals("SEMI-ANNUAL (MAJOR + MINOR)") || PMSName.Equals("SEMI-ANNUAL (GREEN + MINOR)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(6);
                            }

                            else if (PMSName.Equals("QUARTERLY MAJOR (W MAJ SVC)") || PMSName.Equals("QUARTERLY GREEN (W GREEN SVC)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                            }

                            else if (PMSName.Equals("QUARTERLY MINOR (W MAJ SVC)") || PMSName.Equals("QUARTERLY MINOR (W GREEN SVC)") || PMSName.Equals("QUARTERLY MINOR"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                            }

                            else if (PMSName.Equals("QUARTERLY (MAJOR + 3 MINORS)") || PMSName.Equals("QUARTERLY (GREEN + 3 MINORS)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                            }

                            else if (PMSName.Equals("MONTHLY MAJOR (W MAJ SVC)") || PMSName.Equals("MONTHLY GREEN (W GREEN SVC)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service9type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service10type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service11type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service12type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(1).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(2).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(4).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(5).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(7).ToString("yyyy-MM");
                                objPMS["bolt_service9datetext"] = serviceStartDate.AddMonths(8).ToString("yyyy-MM");
                                objPMS["bolt_service10datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service11datetext"] = serviceStartDate.AddMonths(10).ToString("yyyy-MM");
                                objPMS["bolt_service12datetext"] = serviceStartDate.AddMonths(11).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(1);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(2);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(4);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(5);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(7);
                                objPMS["bolt_service9date"] = serviceStartDate.AddMonths(8);
                                objPMS["bolt_service10date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service11date"] = serviceStartDate.AddMonths(10);
                                objPMS["bolt_service12date"] = serviceStartDate.AddMonths(11);
                            }

                            else if (PMSName.Equals("MONTHLY MINOR (W MAJ SVC)") || PMSName.Equals("MONTHLY MINOR (W GREEN SVC)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service9type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service10type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service11type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service12type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(1).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(2).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(4).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(5).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(7).ToString("yyyy-MM");
                                objPMS["bolt_service9datetext"] = serviceStartDate.AddMonths(8).ToString("yyyy-MM");
                                objPMS["bolt_service10datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service11datetext"] = serviceStartDate.AddMonths(10).ToString("yyyy-MM");
                                objPMS["bolt_service12datetext"] = serviceStartDate.AddMonths(11).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(1);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(2);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(4);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(5);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(7);
                                objPMS["bolt_service9date"] = serviceStartDate.AddMonths(8);
                                objPMS["bolt_service10date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service11date"] = serviceStartDate.AddMonths(10);
                                objPMS["bolt_service12date"] = serviceStartDate.AddMonths(11);
                            }

                            else if (PMSName.Equals("MONTHLY (MAJOR + 11 MINORS)") || PMSName.Equals("MONTHLY (GREEN + 11 MINORS)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service9type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service10type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service11type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service12type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(1).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(2).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(4).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(5).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(7).ToString("yyyy-MM");
                                objPMS["bolt_service9datetext"] = serviceStartDate.AddMonths(8).ToString("yyyy-MM");
                                objPMS["bolt_service10datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service11datetext"] = serviceStartDate.AddMonths(10).ToString("yyyy-MM");
                                objPMS["bolt_service12datetext"] = serviceStartDate.AddMonths(11).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(1);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(2);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(4);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(5);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(7);
                                objPMS["bolt_service9date"] = serviceStartDate.AddMonths(8);
                                objPMS["bolt_service10date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service11date"] = serviceStartDate.AddMonths(10);
                                objPMS["bolt_service12date"] = serviceStartDate.AddMonths(11);
                            }

                            else if (PMSName.Equals("2 YEAR QTR (1 MAJOR, 1 GREEN, 6 MINORS)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890002);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(12).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(15).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(18).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(21).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(12);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(15);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(18);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(21);
                            }

                            else if (PMSName.Equals("2 YEAR QTR MAJOR"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(12).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(15).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(18).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(21).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(12);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(15);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(18);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(21);
                            }

                            else if (PMSName.Equals("2 YEAR QTR GREEN"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890002);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890002);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(12).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(15).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(18).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(21).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(12);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(15);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(18);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(21);
                            }

                            else if (PMSName.Equals("2 YEAR QTR MINOR"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(12).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(15).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(18).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(21).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(12);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(15);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(18);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(21);
                            }

                            else if (PMSName.Equals("2 YEAR MON. (1 MAJ, 1 GREEN, 22 MINORS)"))
                            {
                                objPMS["bolt_service1type"] = new OptionSetValue(454890000);
                                objPMS["bolt_service2type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service3type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service4type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service5type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service6type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service7type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service8type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service9type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service10type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service11type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service12type"] = new OptionSetValue(454890001);
                                objPMS["bolt_service1datetext"] = serviceStartDate.ToString("yyyy-MM");
                                objPMS["bolt_service2datetext"] = serviceStartDate.AddMonths(1).ToString("yyyy-MM");
                                objPMS["bolt_service3datetext"] = serviceStartDate.AddMonths(2).ToString("yyyy-MM");
                                objPMS["bolt_service4datetext"] = serviceStartDate.AddMonths(3).ToString("yyyy-MM");
                                objPMS["bolt_service5datetext"] = serviceStartDate.AddMonths(4).ToString("yyyy-MM");
                                objPMS["bolt_service6datetext"] = serviceStartDate.AddMonths(5).ToString("yyyy-MM");
                                objPMS["bolt_service7datetext"] = serviceStartDate.AddMonths(6).ToString("yyyy-MM");
                                objPMS["bolt_service8datetext"] = serviceStartDate.AddMonths(7).ToString("yyyy-MM");
                                objPMS["bolt_service9datetext"] = serviceStartDate.AddMonths(8).ToString("yyyy-MM");
                                objPMS["bolt_service10datetext"] = serviceStartDate.AddMonths(9).ToString("yyyy-MM");
                                objPMS["bolt_service11datetext"] = serviceStartDate.AddMonths(10).ToString("yyyy-MM");
                                objPMS["bolt_service12datetext"] = serviceStartDate.AddMonths(11).ToString("yyyy-MM");
                                objPMS["bolt_service1date"] = serviceStartDate;
                                objPMS["bolt_service2date"] = serviceStartDate.AddMonths(1);
                                objPMS["bolt_service3date"] = serviceStartDate.AddMonths(2);
                                objPMS["bolt_service4date"] = serviceStartDate.AddMonths(3);
                                objPMS["bolt_service5date"] = serviceStartDate.AddMonths(4);
                                objPMS["bolt_service6date"] = serviceStartDate.AddMonths(5);
                                objPMS["bolt_service7date"] = serviceStartDate.AddMonths(6);
                                objPMS["bolt_service8date"] = serviceStartDate.AddMonths(7);
                                objPMS["bolt_service9date"] = serviceStartDate.AddMonths(8);
                                objPMS["bolt_service10date"] = serviceStartDate.AddMonths(9);
                                objPMS["bolt_service11date"] = serviceStartDate.AddMonths(10);
                                objPMS["bolt_service12date"] = serviceStartDate.AddMonths(11);
                            }

                            service.Update(objPMS);
                        }
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("UpdateServiceDates Plugin", ex.ToString());
                        throw;
                    }
                }
            }
        }
    }
}
