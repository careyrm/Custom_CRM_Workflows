using System;
using System.Linq;
using System.Activities;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using CRMToolkitV2.Common;
using Microsoft.Crm.Sdk.Messages;

namespace CRMToolkitV2.CoreWorkflows.Base
{
    public abstract class AddContractLineWorkflowBase : CrmWorkflowBase
    {
        #region Inputs

        [Input("Svc Contract Nbr")]
        public InArgument<string> ContractNbr { get; set; }

        [Input("Start Date")]
        public InArgument<DateTime> StartDate { get; set; }

        [Input("End Date")]
        public InArgument<DateTime> EndDate { get; set; }

        [Input("Bill To Customer")]
        [ReferenceTarget("account")]
        public InArgument<EntityReference> BillToCustomer { get; set; }    
        
        #endregion Inputs

        protected abstract string ContractLineEntityName { get; }
        protected abstract string ParentEntityLookupFieldName { get; }
        protected abstract EntityReference ParentEntity { get; }
        protected abstract void ProcessAdditionalFields(ref Entity record);
        
        protected override void ExecuteWorkflowLogic()
        {
            var queryBids = new QueryExpression("opportunity")
            {
                ColumnSet = new ColumnSet("opportunityid", "selectedlevel", "name", "opportunitynbr", "siteid","siteguid", "opportunityguid", "contractamount", "firstcontractlinecreated"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };

            queryBids.Criteria.AddCondition("ismasterbid", ConditionOperator.Equal, "0");
            queryBids.Criteria.AddCondition("firstcontractlinecreated", ConditionOperator.Equal, "0");
            queryBids.Criteria.AddCondition("svccontractnbr", ConditionOperator.Equal, ContractNbr.Get(Context.ExecutionContext).ToString());
            EntityCollection bids = Context.UserService.RetrieveMultiple(queryBids);
            int bidCount = bids.Entities.Count();
            WriteToLog("Bids Related to Service Contract: " + bidCount.ToString());

            foreach (var bid in bids.Entities)
            {
                var bidTitle = bid.GetAttributeValue<string>("name");
                decimal bidPrice;
                bool convertedPrice = decimal.TryParse(bid.FormattedValues["contractamount"], System.Globalization.NumberStyles.Currency, System.Globalization.CultureInfo.CurrentCulture.NumberFormat, out bidPrice);
                var bidPriceType = bidPrice.GetType();
                var bidLevel = bid.FormattedValues["selectedlevel"];
                var bidNumber = bid.GetAttributeValue<string>("opportunitynbr");
                var siteGUID = bid.GetAttributeValue<string>("siteguid");
                bool createContractLine = true;

                //Check for required fields from the child bid
                if (siteGUID == null)
                {
                    WriteToLog("Missing site guid for child bid");
                    createContractLine = false;
                }
                if (bidNumber == null)
                {
                    WriteToLog("Missing Child bid record");
                    createContractLine = false;
                }

                //if all the required fields are populated then create the contract line
                if (createContractLine)
                {
                    WriteToLog("bidNumber = " + bidNumber.ToString());
                    WriteToLog("bidTitle = " + bidTitle.ToString());
                    WriteToLog("bidPrice = " + bidPrice.ToString());
                    WriteToLog("bidPrice Type =  " + bidPriceType.ToString());
                    WriteToLog("bidLevel = " + bidLevel.ToString());


                    var queryProduct = new QueryExpression("product")
                    {
                        ColumnSet = new ColumnSet("productid", "name"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };

                    queryProduct.Criteria.AddCondition("name", ConditionOperator.Equal, bidLevel.ToString());
                    EntityCollection products = Context.UserService.RetrieveMultiple(queryProduct);
                    WriteToLog("Product Count: " + products.TotalRecordCount.ToString());

                    var contractLine = new Entity(ContractLineEntityName)
                    {
                        [ParentEntityLookupFieldName] = ParentEntity,
                        ["title"] = bidTitle,
                        ["activeon"] = StartDate.Get(Context.ExecutionContext),
                        ["expireson"] = EndDate.Get(Context.ExecutionContext),
                        ["customerid"] = BillToCustomer.Get(Context.ExecutionContext),
                        ["price"] = new Money(bidPrice)
                    };

                    WriteToLog("Adding Product");

                    ProcessAdditionalFields(ref contractLine);

                    products.Entities.ToList().ForEach(product =>
                    {
                        Guid productid = products.Entities[0].Id;
                        WriteToLog("product GUID is: " + productid.ToString());
                        EntityReference relatedProduct = new EntityReference("product", productid);
                        contractLine["productid"] = relatedProduct;
                    });


                    Context.UserService.Create(contractLine);

                    Guid contractlineid = new Guid(contractLine.Id.ToString());
                    WriteToLog("Contract Line GUID: " + contractlineid.ToString());
                    contractLine["contractlineguid"] = contractlineid.ToString();

                    //Get the child bid record to make the relationship from the contract line back to the bid
                    var queryChildBid = new QueryExpression("opportunity")
                    {
                        ColumnSet = new ColumnSet("opportunityid", "opportunitynbr"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };

                    queryChildBid.Criteria.AddCondition("opportunitynbr", ConditionOperator.Equal, bidNumber.ToString());
                    EntityCollection childBids = Context.UserService.RetrieveMultiple(queryChildBid);
                    childBids.Entities.ToList().ForEach(childbid =>
                    {
                        Guid opportunityid = childBids.Entities[0].Id;
                        WriteToLog("Adding Related Child Bid");
                        EntityReference relatedBid = new EntityReference("opportunity", opportunityid);
                        contractLine["nar_relatedbidid"] = relatedBid;
                        WriteToLog("related child bid GUID = " + opportunityid.ToString());
                    });

                    //Relate the contract line to the site account
                    WriteToLog("Adding Related Site");

                    WriteToLog("Site Guid is: " + siteGUID.ToString());
                    Guid siteid = new Guid(siteGUID);
                    EntityReference relatedSite = new EntityReference("account", siteid);
                    contractLine["siteid"] = relatedSite;

                    Context.UserService.Update(contractLine);
                }

            }
        }
    }
}
