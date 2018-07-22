using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace msdyncrmWorkflowTools
{
    public class CompareMultiSelectOptionSet : CodeActivity
    {
        [Output("Result")]
        public OutArgument<bool> Result { get; set; }

        [RequiredArgument]
        [Input("Values")]
        [Default("126530000;126530001")]
        public InArgument<string> Values { get; set; }

        [RequiredArgument]
        [Input("Condition")]
        [Default("And, Or, Null")]
        public InArgument<string> Condition { get; set; }

        [RequiredArgument]
        [Input("AttributeName")]
        public InArgument<string> AttributeName { get; set; }


        protected override void Execute(CodeActivityContext executionContext)
        {
            #region "Load CRM Service from context"

            Common objCommon = new Common(executionContext);
            objCommon.tracingService.Trace("Load CRM Service from context --- OK");

            #endregion

            try
            {
                var values = Values.Get(executionContext).Split(';');
                var attributeName = AttributeName.Get(executionContext).ToLower();
                var condition = Condition.Get(executionContext).ToLower();

                // Add on Trace
                objCommon.tracingService.Trace("Values: {0}", Values.Get(executionContext));
                objCommon.tracingService.Trace("attributeName: {0}", attributeName);
                objCommon.tracingService.Trace("condition: {0}", condition);

                var entity = objCommon.service.Retrieve(objCommon.context.PrimaryEntityName,
                                                        objCommon.context.PrimaryEntityId,
                                                        new ColumnSet(attributeName));

                var optionSetValueCollection = entity.GetAttributeValue<OptionSetValueCollection>(attributeName);
                var result = GetConditionResult(condition, values, optionSetValueCollection);


                Result.Set(executionContext, result);
            }
            catch (Exception ex)
            {
                objCommon.tracingService.Trace("Message: {0} \nStackTrace: {1}", ex.Message, ex.StackTrace);
                throw;
            }
        }

        private bool GetConditionResult(string condition, string[] values, OptionSetValueCollection optionSetValueCollection)
        {
            if (optionSetValueCollection == null) optionSetValueCollection = new OptionSetValueCollection();

            switch (condition)
            {
                case "and":
                    return (from o in optionSetValueCollection
                            join v in values on o.Value.ToString() equals v
                            select o).Count() == values.Count();

                case "or":
                    return (from o in optionSetValueCollection
                            join v in values on o.Value.ToString() equals v
                            select o).Count() > 0;

                case "null":
                    return optionSetValueCollection.Count == 0;

                default:
                    throw new ArgumentException("Condition is not found!");
            }
        }
    }
}