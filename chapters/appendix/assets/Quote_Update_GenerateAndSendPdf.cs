using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using UtilityFlowDevSolution.Classes;

namespace UtilityFlowDevSolution {
  public class Quote_Update_GenerateAndSendPdf : Plugin {
    public Quote_Update_GenerateAndSendPdf() : base(typeof(Quote_Update_GenerateAndSendPdf)) {
      RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(Plugin.PipelineStage.PostOperation, "Update", "quote", GenerateAndSendPdf));
    }

    private void GenerateAndSendPdf(LocalPluginContext context) {
      //Extract the tracing service for use in debugging sandboxed plug-ins.
      ITracingService tracingService = context.TracingService;

      // Obtain the target entity from the input parameters.
      Entity entity = context.GetDefaultPostImage();

      // Verify that the target entity represents a quote.
      // If not, this plug-in was not registered correctly.
      if (entity.LogicalName != "quote")
        return;
      // Verify that target statecode attribute exists (it means that it has been updated)
      if (!entity.Contains("statecode"))
        return;
      // Verify that the target statecode has been updated to "Active"
      if (((OptionSetValue)entity["statecode"]).Value != new OptionSetValue(1).Value)
        return;

      try {
        tracingService.Trace("Plugin triggered");
        IOrganizationService service = context.OrganizationService;

        Guid documentTemplateId = getTemplateId(service);

        // Get quote that triggered the plugin
        Entity quote = service.Retrieve("quote", entity.Id, new ColumnSet(true));

        String b64File = generateEncodedPdf(service, quote, documentTemplateId);
        String attachmentName = "OffertaPersonalizzata.pdf";

        //tracingService.Trace(b64File);

        // create email
        Entity email = generateEmail(service, quote);
        Guid emailId = service.Create(email);

        Guid attachmentId = createAttachment(service, b64File, attachmentName, emailId);

        SendEmailRequest sendEmailReq = new SendEmailRequest {
          EmailId = emailId,
          IssueSend = true,
        };

        SendEmailResponse sendEmailResp = (SendEmailResponse)service.Execute(sendEmailReq);
      } catch (Exception ex) {
        tracingService.Trace("Lead_QualifyLead Plugin: {0}", ex.ToString());
        throw new InvalidPluginExecutionException("An error occurred in the Quote_Update_GenerateAndSendPdf plug-in.", ex);
      }
    }

    private Guid getTemplateId(IOrganizationService service) {
      Guid documentTemplateId = new Guid();
      QueryExpression templateQuery = new QueryExpression("documenttemplate");
      FilterExpression filter = new FilterExpression(LogicalOperator.Or);
      filter.AddCondition("name", ConditionOperator.Like, "Riepilogo Offerta");
      filter.AddCondition("name", ConditionOperator.Like, "Quote Summary");
      templateQuery.Criteria.AddFilter(filter);
      templateQuery.ColumnSet.AddColumns("name", "documenttemplateid", "content");

      EntityCollection resultsTemplate = service.RetrieveMultiple(templateQuery);

      foreach (Entity template in resultsTemplate.Entities) {
        if (template["name"].ToString().Contains("Riepilogo")) {
          return template.Id;
        }
        if (template["name"].ToString() == "Quote Summary") {
          documentTemplateId = template.Id;
        }
      }

      if (documentTemplateId == Guid.Empty) {
        throw new Exception("Template not found.");
      }
      return documentTemplateId;
    }

    private String generateEncodedPdf(IOrganizationService service, Entity quote, Guid documentTemplateId) {
      // Request to create pdf
      OrganizationRequest createPdfRequest = new OrganizationRequest("ExportPdfDocument");
      createPdfRequest["EntityTypeCode"] = 1084; // Quote typecode
      createPdfRequest["SelectedTemplate"] = new EntityReference("documenttemplate", documentTemplateId);
      createPdfRequest["SelectedRecords"] = "[\"{" + quote.Id.ToString() + "}\"]";

      OrganizationResponse createPdfResponse = (OrganizationResponse)service.Execute(createPdfRequest);

      return Convert.ToBase64String((byte[])createPdfResponse["PdfFile"]);
    }

    private Entity generateEmail(IOrganizationService service, Entity quote) {
      Entity email = generateEmailFromTemplate(service, quote);
      return completeEmail(quote, email);
    }

    private Entity generateEmailFromTemplate(IOrganizationService service, Entity quote) {
      //Create a query expression to get necessary email quote template
      QueryExpression queryBuildInTemplates = new QueryExpression {
        EntityName = "template",
        ColumnSet = new ColumnSet("templateid", "title"),
        Criteria = new FilterExpression()
      };
      queryBuildInTemplates.Criteria.AddCondition("title", ConditionOperator.Equal, "Template Offerta Email");
      EntityCollection templateEntityCollection = service.RetrieveMultiple(queryBuildInTemplates);

      Guid emailTemplateId = new Guid();

      if (templateEntityCollection.Entities.Count > 0) {
        emailTemplateId = (Guid)templateEntityCollection.Entities[0].Attributes["templateid"];
      } else {
        throw new ArgumentException("Standard Email Templates are missing");
      }

      // Create the request
      InstantiateTemplateRequest emailUsingTemplateReq = new InstantiateTemplateRequest {
        // Use a built-in Email Template of type "quote".
        TemplateId = emailTemplateId,

        // The regarding Id is required, and must be of the same type as the Email Template.
        ObjectId = quote.Id,
        ObjectType = "quote"
      };

      InstantiateTemplateResponse emailUsingTemplateResp = (InstantiateTemplateResponse)service.Execute(emailUsingTemplateReq);
      return emailUsingTemplateResp.EntityCollection[0];
    }

    private Entity completeEmail(Entity quote, Entity email) {
      Guid clientId = ((EntityReference)quote.Attributes["customerid"]).Id;
      Guid senderId = ((EntityReference)quote.Attributes["ownerid"]).Id; // Quote owner is the sender

      // Create the 'From:' activity party for the email
      Entity fromParty = new Entity("activityparty");
      fromParty["partyid"] = new EntityReference("systemuser", senderId);

      // Create the 'To:' activity party for the email
      Entity toParty = new Entity("activityparty");
      toParty["partyid"] = new EntityReference("account", clientId);

      email["to"] = new Entity[] { toParty };
      email["from"] = new Entity[] { fromParty };
      email["directioncode"] = true;
      email["regardingobjectid"] = new EntityReference("quote", quote.Id);

      return email;
    }

    private Guid createAttachment(IOrganizationService service, String b64File, String attachmentName, Guid emailId) {
      // Create attachment
      Entity attachment = new Entity("activitymimeattachment");
      attachment["body"] = b64File;
      attachment["filename"] = attachmentName;
      attachment["objectid"] = new EntityReference("email", emailId);
      attachment["objecttypecode"] = "email";
      return service.Create(attachment);
    }
  }
}
