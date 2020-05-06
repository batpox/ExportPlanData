using SimioAPI;
using SimioAPI.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportPlanData
{
    public class OnSaveInstance : IDisposable
    {
        readonly IModelHelperContext _context;

        public OnSaveInstance(IModelHelperContext context)
        {
            _context = context;

            // The context contains various model events you can subscribe to. You will usually
            //  want to unsubscribe from the event in the implementation of Dispose()
            _context.ModelSaved += _context_ModelSaved;

        }

        /// <summary>
        /// This is the handler for the "Model Saved" event, this is the bulk of the "helping" logic
        /// </summary>
        /// <param name="args"></param>
        private void _context_ModelSaved(IModelSavedArgs args)
        {
            try
            {
                string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                DataTable dtUsageLog = null;

                dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.MaterialUsageLog);
                dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.ResourceUsageLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.ResourceCapacityLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.ResourceInfoLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.ResourceStateLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.ConstraintLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.StateObservationLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.TallyObservationLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.TaskLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.TaskStateLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                //dtUsageLog = SimioAPIHelpers.LogHelpers.ConvertLogToDataTable(_context.Model.Plan.TransporterUsageLog);
                //dtUsageLog.WriteXml(Path.Combine(myDocs, dtUsageLog.TableName + ".xml"));

                int tableCount = 0;
                foreach ( ITable simioTable in _context.Model.Tables)
                {
                    tableCount++;
                    DataTable dtTable = SimioAPIHelpers.TableHelpers.ConvertSimioTableToDataTable(simioTable);
                    if (dtTable != null)
                    {
                        string tablename = simioTable.Name + ".xml";
                        string outputFilename = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), tablename);
                        dtTable.WriteXml(outputFilename);
                    }
                }

            }
            catch (Exception ex)
            {
                Alert(ex.Message);
            }
        }

        public void Dispose()
        {
            

            // Unsubscribing from the events here, as we no longer need to listen to them,
            //  Dispose() indicates the addin has been "unloaded"
            _context.ModelSaved -= _context_ModelSaved;
        }

        private void Alert(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }


    }
}
