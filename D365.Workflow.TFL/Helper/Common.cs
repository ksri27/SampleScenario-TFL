using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// Common functions can be used at multiple places
/// </summary>
namespace D365.Workflow.TFL.Helper
{
    public class Common
    {
        public static string caseEmailAttFetchXml = @"<fetch>
                              <entity name='activitymimeattachment'>
                                <attribute name='subject' />
                                <attribute name='attachmentid' />
                                <link-entity name='activitypointer' from='activityid' to='objectid' link-type='inner'>
                                  <link-entity name='incident' from='incidentid' to='regardingobjectid' link-type='inner'>
                                    <filter>
                                      <condition attribute='incidentid' operator='eq' value='{0}' uitype='incident' />
                                    </filter>
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>";

        public static string contactEmaiAttFetchXml = @"<fetch>
                              <entity name='activitymimeattachment'>
                                <attribute name='subject' />
                                <attribute name='attachmentid' />
                                <link-entity name='activitypointer' from='activityid' to='objectid' link-type='inner'>
                                  <link-entity name='contact' from='contactid' to='regardingobjectid' link-type='inner'>
                                    <filter>
                                      <condition attribute='contactid' operator='eq' value='{0}' uitype='contact' />
                                    </filter>
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>";

        public static string notesAttFetchXml = @"<fetch>
                                      <entity name='annotation'>
                                        <attribute name='filename' />
                                        <attribute name='mimetype' />
                                        <attribute name='documentbody' />
                                        <filter type='and'>
                                          <condition attribute='isdocument' operator='eq' value='1' />
                                          <filter type='or'>
                                            <condition attribute='objectid' operator='eq' value='{0}' uitype='contact' />
                                            <condition attribute='objectid' operator='eq' value='{1}' uitype='incident' />
                                          </filter>
                                        </filter>
                                      </entity>
                                    </fetch>";

        public static string teamMembFetchXml = @"<fetch>
                                         <entity name='systemuser'>
                                           <attribute name='fullname' />
                                           <attribute name='systemuserid' />
                                           <order attribute='fullname' descending='false' />
                                           <link-entity name='teammembership' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                                             <link-entity name='team' from='teamid' to='teamid' alias='ac'>
                                               <filter type='and'>
                                                 <condition attribute='teamid' operator='eq' uitype='team' value='{0}' />
                                               </filter>
                                             </link-entity>
                                           </link-entity>
                                         </entity>
                                       </fetch>";

        /// <summary>
        /// Create Attachments to Email from other activities
        /// </summary>
        /// <param name="emailId">Email Guid</param>
        /// <param name="emailAttachmentsCollections">Email attachment collection</param>
        /// <param name="service">Organization service</param>
        public static void CreateAttachmentsfromEmail(Guid emailId, EntityCollection emailAttachmentsCollections, IOrganizationService service)
        {
            try
            {
                if (emailAttachmentsCollections.Entities.Count > 0)
                {
                    foreach (Entity emailAttachment in emailAttachmentsCollections.Entities)
                    {
                        Entity attachment = new Entity("activitymimeattachment");
                        if (emailAttachment.Attributes.Contains("subject"))
                        {
                            attachment["subject"] = emailAttachment.GetAttributeValue<string>("subject");
                        }
                        attachment["objectid"] = new EntityReference("email", emailId);
                        attachment["objecttypecode"] = "email";
                        attachment["attachmentid"] = emailAttachment.GetAttributeValue<EntityReference>("attachmentid");
                        service.Create(attachment);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Create Attachments to Email from notes
        /// </summary>
        /// <param name="emailId">Email Guid</param>
        /// <param name="notesAttachCollections">notes entity collection</param>
        /// <param name="service">Organization service</param>
        public static void CreateAttachmentsfromAnnotation(Guid emailId, EntityCollection notesAttachCollections, IOrganizationService service)
        {
            try
            {
                if (notesAttachCollections.Entities.Count > 0)
                {
                    foreach (Entity annotationAttachment in notesAttachCollections.Entities)
                    {
                        Entity attachment = new Entity("activitymimeattachment");
                        attachment["objectid"] = new EntityReference("email", emailId);
                        attachment["objecttypecode"] = "email";
                        attachment["subject"] = annotationAttachment.GetAttributeValue<string>("filename");
                        attachment["filename"] = annotationAttachment.GetAttributeValue<string>("filename"); ;
                        attachment["body"] = annotationAttachment.GetAttributeValue<string>("documentbody");
                        attachment["mimetype"] = annotationAttachment.GetAttributeValue<string>("mimetype");
                        service.Create(attachment);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Add team members to Access team
        /// </summary>
        /// <param name="recordReference">Record EntityReference</param>
        /// <param name="teamReference">Team EntityReference</param>
        /// <param name="accessTeamId">Access team Id</param>
        /// <param name="service">Organization service</param>
        public static void AddTeamtoAccessTeam(EntityReference recordReference, EntityReference teamReference, Guid accessTeamId, IOrganizationService service)
        {
            try
            {
                EntityCollection teamMembCollections = service.RetrieveMultiple(new FetchExpression(string.Format(Common.teamMembFetchXml, teamReference.Id)));
                if (teamMembCollections.Entities.Count > 0)
                {
                    foreach (Entity teamMember in teamMembCollections.Entities)
                    {
                        AddUsertoAccessTeam(recordReference, teamMember.Id, accessTeamId, service);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Add user to access team
        /// </summary>
        /// <param name="recordReference">Record EntityReference</param>
        /// <param name="userId">User Guid</param>
        /// <param name="accessTeamId">Access team Id</param>
        /// <param name="service">Organization service</param>
        public static void AddUsertoAccessTeam(EntityReference recordReference, Guid userId, Guid accessTeamId, IOrganizationService service)
        {
            try
            {
                AddUserToRecordTeamRequest addUserRequestOwner = new AddUserToRecordTeamRequest()
                {
                    Record = recordReference,
                    SystemUserId = userId,
                    TeamTemplateId = accessTeamId,
                };
                service.Execute(addUserRequestOwner);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Assign record to user or team
        /// </summary>
        /// <param name="recordAssignee">Assignee EntityReference</param>
        /// <param name="recordReference">Record EntityReference</param>
        /// <param name="service">Organization service</param>
        public static void AssignRecord(EntityReference recordAssignee, EntityReference recordReference, IOrganizationService service)
        {
            try
            {
                AssignRequest assignOwner = new AssignRequest
                {
                    Assignee = recordAssignee,
                    Target = recordReference
                };
                service.Execute(assignOwner);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Send draft email
        /// </summary>
        /// <param name="recordId">Email Guid</param>
        /// <param name="service">Organization Guid</param>
        public static void SendEmail(Guid recordId, IOrganizationService service)
        {
            try
            {
                SendEmailRequest sendEmailRequest = new SendEmailRequest
                {
                    EmailId = recordId,
                    TrackingToken = "",
                    IssueSend = true
                };
                SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
