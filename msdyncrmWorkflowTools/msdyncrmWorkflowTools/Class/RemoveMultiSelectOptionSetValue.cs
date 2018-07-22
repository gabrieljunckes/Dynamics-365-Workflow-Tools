using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace msdyncrmWorkflowTools
{
    public class RemoveMultiSelectOptionSetValue : CodeActivity
    {
        [RequiredArgument]
        [Input("Values")]
        [Default("126530000;126530001")]
        public InArgument<string> Values { get; set; }

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

                // Add on Trace
                objCommon.tracingService.Trace("Values: {0}", Values.Get(executionContext));
                objCommon.tracingService.Trace("attributeName: {0}", attributeName);

                var entity = objCommon.service.Retrieve(objCommon.context.PrimaryEntityName,
                                                        objCommon.context.PrimaryEntityId,
                                                        new ColumnSet(attributeName));

                var optionSetValueCollection = entity.GetAttributeValue<OptionSetValueCollection>(attributeName);
                var newEntity = new Entity(objCommon.context.PrimaryEntityName, objCommon.context.PrimaryEntityId);

                newEntity.Attributes[attributeName] = RemoveValuesInOptionSetValueCollection(values, optionSetValueCollection);

                objCommon.tracingService.Trace("Update record");
                objCommon.service.Update(newEntity);
            }
            catch (Exception ex)
            {
                objCommon.tracingService.Trace("Message: {0} \nStackTrace: {1}", ex.Message, ex.StackTrace);
                throw;
            }
        }

        private OptionSetValueCollection RemoveValuesInOptionSetValueCollection(string[] values, OptionSetValueCollection optionSetValueCollection)
        {
            var newOptionSetValueCollection = new OptionSetValueCollection();

            foreach (var option in optionSetValueCollection)
            {
                if (!values.Any(x => x == option.Value.ToString()))
                {
                    newOptionSetValueCollection.Add(option);
                }
            }

            return newOptionSetValueCollection;
        }
    }
}