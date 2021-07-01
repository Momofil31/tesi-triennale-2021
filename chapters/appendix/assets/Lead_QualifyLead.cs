using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using UtilityFlowDevSolution.Classes;

namespace UtilityFlowDevSolution {
  /* Nel costruttore mettere tutti gli step del plugin 
  * costruttore public nomeclasse () : base(typeof(nomeclasse)) { 
  *   RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(Plugin.PipelineStage.postOperation, "Create", "lead",  ))
  * }
  * 
  */
  public class Lead_QualifyLead : Plugin {
    public Lead_QualifyLead() : base(typeof(Lead_QualifyLead)) {
      RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>(Plugin.PipelineStage.PostOperation, "Create", "lead", QualifyLead));
    }

    private void QualifyLead(LocalPluginContext context) {
      //Extract the tracing service for use in debugging sandboxed plug-ins.
      ITracingService tracingService = context.TracingService;

      // Obtain the target entity from the input parameters.
      Entity entity = context.GetDefaultPostImage();

      // Verify that the target entity represents a lead.
      // If not, this plug-in was not registered correctly.
      if (entity.LogicalName != "lead")
        return;

      try {
        IOrganizationService service = context.OrganizationService;
        Guid accountId = new Guid();
        Guid contactId = new Guid();
        Guid opportunityId = new Guid();


        // Get companyname to check if account exists
        String companyName = entity.GetAttributeValue<string>("companyname");
        String partitaIva = entity.GetAttributeValue<string>("rfu_partitaiva");
        Boolean accountExists = false;


        // Check if the account already exist 
        QueryExpression accountQuery = new QueryExpression("account");
        accountQuery.NoLock = true;
        accountQuery.ColumnSet = new ColumnSet("accountid");
        accountQuery.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Like, companyName));
        if (partitaIva != null) accountQuery.Criteria.AddCondition(new ConditionExpression("accountnumber", ConditionOperator.Like, partitaIva));
        accountQuery.PageInfo.ReturnTotalRecordCount = true;
        EntityCollection resultsAccount = service.RetrieveMultiple(accountQuery);

        // Get email to check if contact exists
        String email = entity.GetAttributeValue<string>("emailaddress1");
        String mobile = entity.GetAttributeValue<string>("mobilephone");
        Boolean contactExists = false;

        // Check if the contact already exist 
        QueryExpression contactQuery = new QueryExpression("contact");
        contactQuery.NoLock = true;
        contactQuery.ColumnSet = new ColumnSet("contactid");
        contactQuery.Criteria.AddCondition(new ConditionExpression("emailaddress1", ConditionOperator.Like, email));
        if (mobile != null) contactQuery.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Like, mobile));
        contactQuery.PageInfo.ReturnTotalRecordCount = true;
        EntityCollection resultsContact = service.RetrieveMultiple(contactQuery);

        if (resultsAccount.TotalRecordCount > 0) {
          accountExists = true;
          foreach (Entity item in resultsAccount.Entities) {
            accountId = item.Id; // get account id
            break;
          }
        }
        if (resultsContact.TotalRecordCount > 0) {
          contactExists = true;
          foreach (Entity item in resultsContact.Entities) {
            contactId = item.Id; // get contact id
            break;
          }
        }

        // Retrieve the organization's base currency ID for setting the
        // transaction currency of the opportunity.
        var query = new QueryExpression("organization");
        query.ColumnSet = new ColumnSet("basecurrencyid");
        var result = service.RetrieveMultiple(query);
        var currencyId = (EntityReference)result.Entities[0]["basecurrencyid"];

        QualifyLeadRequest qualify = new QualifyLeadRequest {
          CreateOpportunity = true,
          OpportunityCurrencyId = currencyId,
          LeadId = new EntityReference(entity.LogicalName, entity.Id),
          Status = new OptionSetValue(OptionSet.Lead.StatusCode.Qualified)
        };

        if (contactExists && !accountExists) {
          // Create account and use existing contact as OpportunityCustomerId
          qualify.CreateAccount = true;
          qualify.OpportunityCustomerId = new EntityReference("contact", contactId);
        } else if (!contactExists && accountExists) {
          // Create contact and use existing account as OpportunityCustomerId
          qualify.CreateContact = true;
          qualify.OpportunityCustomerId = new EntityReference("account", accountId);
        } else if (contactExists && accountExists) {
          // Use existing account as OpportunityCustomerId
          qualify.OpportunityCustomerId = new EntityReference("account", accountId);
          //if contact is not associated to account 
        } else {
          qualify.CreateAccount = true;
          qualify.CreateContact = true;
        }

        //Disable DuplicateDetection
        qualify.Parameters.Add("SuppressDuplicateDetection", true);

        QualifyLeadResponse qualifyRes = (QualifyLeadResponse)service.Execute(qualify);

        foreach (var resEntity in qualifyRes.CreatedEntities) {
          //NotifyEntityCreated(entity.LogicalName, entity.Id);
          if (resEntity.LogicalName == "account") {
            accountId = resEntity.Id;
          }
          if (resEntity.LogicalName == "contact") {
            contactId = resEntity.Id;
          }
          if (resEntity.LogicalName == "opportunity") {
            opportunityId = resEntity.Id;
            tracingService.Trace("OpportunityID: " + opportunityId);
          }
        }

        // Update account if newly created to add account number
        if (!accountExists) {
          Entity retrievedAccount = new Entity("account", accountId);
          Entity account = new Entity("account");
          account.Id = retrievedAccount.Id;
          account["accountnumber"] = partitaIva;
          service.Update(account);
          tracingService.Trace("Updated account " + accountId);
        }

        String[] fieldsToUpdate = new string[] {
                    "rfu_consumielettgenfeb",
                    "rfu_consumielettmaggiu",
                    "rfu_consumielettmarapr",
                    "rfu_consumielettnovdic",
                    "rfu_consumielettrlugago",
                    "rfu_consumielettsetott",
                    "rfu_consumigasgenfeb",
                    "rfu_consumigaslugago",
                    "rfu_consumigasmaggiu",
                    "rfu_consumigasmarapr",
                    "rfu_consumigasnovdic",
                    "rfu_consumigassetott",
                    "rfu_electricityintendeduse",
                    "rfu_energyrating",
                    "rfu_floorarea",
                    "rfu_gasintendeduse",
                    "rfu_impiantoproprio",
                    "rfu_requestedutility",
                    "rfu_tensioneallacciamento",
                    "rfu_tipocontratto"
                };

        // Add custom field's data to created opportunity
        Entity retrievedOpportunity = new Entity("opportunity", opportunityId);
        tracingService.Trace(retrievedOpportunity.ToString());
        Entity opportunity = new Entity("opportunity");
        opportunity.Id = retrievedOpportunity.Id;

        if (!qualify.CreateContact) {
          opportunity["parentcontactid"] = new EntityReference("contact", contactId);
        }

        foreach (String field in fieldsToUpdate) {
          opportunity[field] = entity.Contains(field) ? entity[field] : null;
        }

        tracingService.Trace("before update opportunity");
        service.Update(opportunity);
        tracingService.Trace("After update opportunity");

      } catch (Exception ex) {
        tracingService.Trace("Lead_QualifyLead Plugin: {0}", ex.ToString());
        throw new InvalidPluginExecutionException("An error occurred in the Lead_QualifyLead plug-in.", ex);
      }
    }
  }
}
