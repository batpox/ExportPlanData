using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SimioAPI;
using SimioAPI.Extensions;

namespace ExportPlanData
{
    class ExportPlanDataAddIn : IModelHelperAddIn
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "ExportPlanData"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.
        /// </summary>
        public string Description
        {
            get { return "Demonstrates the conversion of Plan Data (Logs, Tables, ...) to DataTables and exporting to XML"; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return null; }
        }

        public Guid UniqueID => new Guid("046EA880-68D7-443E-8A56-7328997AF282");

        public ModelAddInEnvironment AddInEnvironment => ModelAddInEnvironment.InteractiveDesktop;

        public IDisposable CreateInstance(IModelHelperContext context)
        {
            return new OnSaveInstance(context);

        }

        public void DefineSchema(IModelHelperAddInSchema schema)
        {

        }

        /// <summary>
        /// Method called when the add-in is run.
        /// </summary>
        public void Execute(IDesignContext context)
        {
            // This example code places some new objects from the Standard Library into the active model of the project.
            if (context.ActiveModel != null)
            {
                // Example of how to place some new fixed objects into the active model.
                // This example code places three new fixed objects: a Source, a Server, and a Sink.
                IIntelligentObjects intelligentObjects = context.ActiveModel.Facility.IntelligentObjects;
                IFixedObject sourceObject = intelligentObjects.CreateObject("Source", new FacilityLocation(-10, 0, -10)) as IFixedObject;
                IFixedObject serverObject = intelligentObjects.CreateObject("Server", new FacilityLocation(0, 0, 0)) as IFixedObject;
                IFixedObject sinkObject = intelligentObjects.CreateObject("Sink", new FacilityLocation(10, 0, 10)) as IFixedObject;

                // Example of how to place some new link objects into the active model (to add network paths between nodes).
                // This example code places two new link objects: a Path connecting the Source 'output' node to the Server 'input' node,
                // and a Path connecting the Server 'output' node to the Sink 'input' node.
                INodeObject sourceOutputNode = sourceObject.Nodes[0];
                INodeObject serverInputNode = serverObject.Nodes[0];
                INodeObject serverOutputNode = serverObject.Nodes[1];
                INodeObject sinkInputNode = sinkObject.Nodes[0];
                ILinkObject pathObject1 = intelligentObjects.CreateLink("Path", sourceOutputNode, serverInputNode, null) as ILinkObject;
                ILinkObject pathObject2 = intelligentObjects.CreateLink("Path", serverOutputNode, sinkInputNode, null) as ILinkObject;

                // Example of how to edit the property of an object.
                // This example code edits the 'ProcessingTime' property of the added Server object.
                serverObject.Properties["ProcessingTime"].Value = "100";
            }
        }

        private void Alert(string message)
        {
            MessageBox.Show(message);
        }

        #endregion
    }
}
