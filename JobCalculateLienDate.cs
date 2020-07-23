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
    public abstract class JobCalculateLienDate : CrmWorkflowBase
    {
        #region Inputs      

        [Input("Job Location State")]
        public InArgument<string> JobAddressState { get; set; }

        [Input("Final Date of Work")]
        public InArgument<DateTime> FinalDateOfWork { get; set; }

        [Output("Result")]
        public OutArgument<DateTime> Result { get; set; }

        #endregion Inputs

        protected abstract string JobEntityName { get; }
        protected abstract void ProcessAdditionalFields(ref Entity record);

        protected override void ExecuteWorkflowLogic()
        {
            var queryLienDeadlines = new QueryExpression("liendeadlines")
            {
                ColumnSet = new ColumnSet("subdeadlinedays"),
                Criteria = new FilterExpression(LogicalOperator.And),
                TopCount = 1
            };
			
			//Query the custom entity LienDeadlines to get the sub deadline days based on the locations state.
            queryLienDeadlines.Criteria.AddCondition("locationstate", ConditionOperator.Equal, JobAddressState.Get(Context.ExecutionContext).ToString());
            EntityCollection lienDeadlines = Context.UserService.RetrieveMultiple(queryLienDeadlines);
            int lienCount = lienDeadlines.Entities.Count();
            WriteToLog("Lien Deadline By State Count: " + lienCount.ToString());

            Int32 subDeadlineDays;
            int.TryParse(lienDeadlines.Entities[0].FormattedValues["subdeadlinedays"], out subDeadlineDays);
            WriteToLog("Lien Sub Deadline Days: " + subDeadlineDays.ToString());

			//Calculate and return the lien deadline date based on the retrieved lien sub deadline days value.
            var finalDateofWork = FinalDateOfWork.Get(Context.ExecutionContext);
            DateTime lienDeadlineDate = finalDateofWork.AddDays(subDeadlineDays);
            WriteToLog("Calculated Lien Deadline Date: " + lienDeadlineDate.ToString());

            Result.Set(Context.ExecutionContext, lienDeadlineDate);

        }

    }
}
