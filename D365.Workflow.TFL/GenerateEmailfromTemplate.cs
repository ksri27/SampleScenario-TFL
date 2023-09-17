using D365.Workflow.TFL.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;

namespace D365.Workflow.TFL
{
    /// <summary>
    /// Custom Workfow to create email and add attachemnts from case, contact and related activites
    /// Send email to teams
    /// Based on condition add members to confidential team and to make the record visible only to specific users
    /// </summary>
    public class GenerateEmailfromTemplate : CodeActivity
    {
        [RequiredArgument]
        [Input("Team Name")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> TeamName { get; set; }

        [RequiredArgument]
        [Input("Contact")]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> Contact { get; set; }

        [RequiredArgument]
        [Input("Escalation Email Template")]
        [ReferenceTarget("template")]
        public InArgument<EntityReference> EscTemplate { get; set; }

        [RequiredArgument]
        [Input("Confidential Email Template")]
        [ReferenceTarget("template")]
        public InArgument<EntityReference> ConTemplate { get; set; }

        [RequiredArgument]
        [Input("Confidential Email?")]
        public InArgument<bool> IsConfidential { get; set; }

        [RequiredArgument]
        [Input("Parent Confidential Team")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> ParentConfidential { get; set; }

        [RequiredArgument]
        [Input("Access Team Id")]
        public InArgument<string> AccessTeam { get; set; }

        public static IOrganizationService _service = null;
        public static string accessTeam = string.Empty;
        public static EntityReference parentConfidentialTeam;
        public static EntityReference teamReference;
        public static EntityReference contactRef;
        public static EntityReference escalationTemplateName;
        public static EntityReference confidentialTemplateName;
        public static EntityReference regardingObjRef;
        public static EntityReference templateName;
        public static Guid userId;
        public static bool isConfidentialMail;

        ITracingService tracingService;
        protected override void Execute(CodeActivityContext executionContext)
        {
            tracingService = executionContext.GetExtension<ITracingService>();
            CRMHelper.getCRMServices(executionContext);
            _service = CRMHelper.OrgService;
            tracingService.Trace("GenerateEmailfromTemplate: Execution Started");
            try
            {
                teamReference = TeamName.Get<EntityReference>(executionContext);
                contactRef = Contact.Get<EntityReference>(executionContext);
                escalationTemplateName = EscTemplate.Get<EntityReference>(executionContext);
                confidentialTemplateName = ConTemplate.Get<EntityReference>(executionContext);
                parentConfidentialTeam = ParentConfidential.Get<EntityReference>(executionContext);
                userId = CRMHelper.WorkflowContext.UserId;
                regardingObjRef = new EntityReference(CRMHelper.WorkflowContext.PrimaryEntityName, CRMHelper.WorkflowContext.PrimaryEntityId);
                templateName = null;
                isConfidentialMail = IsConfidential.Get<bool>(executionContext);
                accessTeam = AccessTeam.Get<string>(executionContext);
                if (isConfidentialMail)
                {
                    templateName = confidentialTemplateName;
                }
                else
                {
                    templateName = escalationTemplateName;
                }
                tracingService.Trace("Email Template Id: " + templateName.Id.ToString());
                SendEmailWithAttachments();

            }
            catch(Exception ex)
            {
                tracingService.Trace(ex.Message.ToString());
                throw new InvalidPluginExecutionException("ExecuteAction Exception:" + ex.Message.ToString());
            }
        }

        public void SendEmailWithAttachments()
        {
            EntityReference queueRef = null;
            Guid templateId = Guid.Empty;

            QueryExpression queryBuildInQueue = new QueryExpression
            {
                EntityName = "queue",
                ColumnSet = new ColumnSet("queueid"),
                Criteria = new FilterExpression()
            };
            queryBuildInQueue.Criteria.AddCondition("ownerid",
                                                        ConditionOperator.Equal, teamReference.Id);
            EntityCollection queueEntityCollection = _service.RetrieveMultiple(queryBuildInQueue);
            if (queueEntityCollection.Entities.Count > 0)
            {
                queueRef = queueEntityCollection.Entities[0].ToEntityReference();
            }
            tracingService.Trace("Queue Id: " + queueRef.Id.ToString());
            templateId = templateName.Id;

            InstantiateTemplateRequest instTemplateReq = new InstantiateTemplateRequest
            {
                TemplateId = templateId,
                ObjectId = regardingObjRef.Id,
                ObjectType = regardingObjRef.LogicalName
            };
            InstantiateTemplateResponse instTemplateResp = (InstantiateTemplateResponse)_service.Execute(instTemplateReq);

            Entity template = instTemplateResp.EntityCollection.Entities[0];

            Entity fromParty = new Entity("activityparty");
            fromParty["partyid"] = new EntityReference("systemuser", userId);
            var listfrom = new List<Entity>()
            {
                fromParty
            };
            Entity toParty = new Entity("activityparty");
            toParty["partyid"] = queueRef;
            var listto = new List<Entity>()
            {
                toParty
            };

            Entity email = new Entity("email");
            email["to"] = new EntityCollection(listto);
            email["from"] = new EntityCollection(listfrom);
            if (template != null && template.Attributes.Contains("description"))
            {
                email["subject"] = template.Attributes["subject"].ToString();//from template
                email["description"] = template.Attributes["description"].ToString();//from template
            }
            email["regardingobjectid"] = regardingObjRef;
            email.Id = _service.Create(email);
            tracingService.Trace("Email Id: " + email.Id.ToString());
            AddAttachmenttoEmail(email.Id);

            if (isConfidentialMail)
            {
                Common.AddTeamtoAccessTeam(new EntityReference("email", email.Id), teamReference, Guid.Parse(accessTeam), _service);
                Common.AddUsertoAccessTeam(new EntityReference("email", email.Id), userId, Guid.Parse(accessTeam), _service);
                Common.AssignRecord(parentConfidentialTeam, new EntityReference("email", email.Id), _service);
                tracingService.Trace("Confidential Team Email");
            }

            Common.SendEmail(email.Id, _service);
        }

        public void AddAttachmenttoEmail(Guid emailguid)
        {
            EntityCollection caseEmailAttachCollections = _service.RetrieveMultiple(new FetchExpression(string.Format(Common.caseEmailAttFetchXml, regardingObjRef.Id)));
            Common.CreateAttachmentsfromEmail(emailguid, caseEmailAttachCollections, _service);


            EntityCollection contactEmailAttachCollections = _service.RetrieveMultiple(new FetchExpression(string.Format(Common.contactEmaiAttFetchXml, contactRef.Id)));
            Common.CreateAttachmentsfromEmail(emailguid, contactEmailAttachCollections, _service);

            EntityCollection notesAttachCollections = _service.RetrieveMultiple(new FetchExpression(string.Format(Common.notesAttFetchXml, contactRef.Id, regardingObjRef.Id)));
            Common.CreateAttachmentsfromAnnotation(emailguid, notesAttachCollections, _service);
            tracingService.Trace("Attachment Added");
        }
    }
}
