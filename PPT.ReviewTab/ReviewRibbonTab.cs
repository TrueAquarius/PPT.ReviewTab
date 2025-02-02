using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPT.ReviewTab.Code.Model;
using PPT.ReviewTab.Code.Util;
using PPT.ReviewTab.WPF;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using static System.Net.Mime.MediaTypeNames;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace PPT.ReviewTab
{
    [ComVisible(true)]
    public class ReviewRibbonTab : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        private ConfigurationManager configurationManager = ConfigurationManager.Instance;

        public ReviewRibbonTab()
        {
            ConfigurationManager.Instance.LoadConfiguration();
            ConfigurationManager.Instance.ConfigurationChanged += OnConfigurationChanged;
            ConfigurationManager.Instance.NotifyConfigurationChanged(null);
        }



        private void OnConfigurationChanged(object sender, EventArgs e)
        {
            if (ribbon == null) return;
            if (sender == null) return;


            int groupNumber = 0;
            foreach(ItemGroup group in ConfigurationManager.Instance.Configuration.ItemGroups)
            {
                ++groupNumber;
                ribbon.InvalidateControl("DynamicGroup_" + groupNumber);

                int itemNumber = 0;
                foreach(Item item in group.Items)
                {

                    ++itemNumber;
                    string itemName = "DynamicItem_" + groupNumber + "_" + itemNumber + "_TXT";
                    ribbon.InvalidateControl(itemName);
                }
            }


            switch(sender.GetType().Name)
            {
                case "Item":
                    break;
                case "ItemGroup":

                    break;

            }
        }





        /// <summary>
        /// Dynamically create the content.
        /// </summary>
        /// <returns></returns>
        public string CreateDynamicGroups()
        {
            // Generate groups dynamically based on configuration
            string groupsXml = string.Empty;

            if (ConfigurationManager.Instance.Configuration == null || ConfigurationManager.Instance.Configuration.ItemGroups == null)
                return groupsXml;

            var groupsXmlBuilder = new System.Text.StringBuilder();
            int groupNumber = 0;
            foreach (var itemGroup in ConfigurationManager.Instance.Configuration.ItemGroups)
            {
                ++groupNumber;
                string groupName = "DynamicGroup_" + groupNumber;
                groupsXmlBuilder.Append($@"
                <group id='{groupName}' label='{itemGroup.Name}'>");

                int itemNumber = 0;
                foreach (var item in itemGroup.Items)
                {
                    ++itemNumber;
                    string itemName = "DynamicItem_" +groupNumber + "_" + itemNumber;
                    string itemXml = string.Empty;

                    itemXml += $"<box     id='{itemName}_BOX' boxStyle='horizontal'>";
                    itemXml += $"<editBox id='{itemName}_TXT' label='{itemNumber}' sizeString='XXXXXXXXXXXXXX' getText='GetEditBoxText' onChange='OnEditBoxChanged' />";
                    itemXml += $"<button  id='{itemName}_BTN' label='Go' onAction='OnButtonAction' />";
                    itemXml += $"</box>";

                   
                    groupsXmlBuilder.Append(itemXml);
                }


                groupsXmlBuilder.Append($"<button id='BTN_{groupNumber}_0_REM' label='remove' onAction='OnRemoveTagAction' />");
                groupsXmlBuilder.Append("</group>");
            }

            groupsXml = groupsXmlBuilder.ToString();

            return groupsXml;
        }



        public string GetEditBoxText(IRibbonControl control)
        {
            Item item = GetItem(control);

            if (item == null || item.Name == null) return "";

            return item.Name;
        }


        public Item GetItem(IRibbonControl control)
        {
            try
            {
                // Provide default or initial text for the editBox
                string[] s = control.Id.Split('_');
                int groupNumber = int.Parse(s[1]);
                int itemNumber = int.Parse(s[2]);

                return ConfigurationManager.Instance.Configuration.ItemGroups[groupNumber - 1].Items[itemNumber - 1];
            }
            catch { }

            return null;
        }

        public ItemGroup GetItemGroup(IRibbonControl control)
        {
            try
            {
                // Provide default or initial text for the editBox
                string[] s = control.Id.Split('_');
                int groupNumber = int.Parse(s[1]);
                //int itemNumber = int.Parse(s[2]);

                return ConfigurationManager.Instance.Configuration.ItemGroups[groupNumber - 1];
            }
            catch { }

            return null;
        }



        public void MoveSlidesToEndClicked(IRibbonControl control)
        {
            var slideMover = new SlideMover(Globals.ThisAddIn.Application);
            slideMover.MoveSelectedSlidesToEnd();
        }





        public void OnRemoveTagAction(IRibbonControl control)
        {
            ItemGroup group = GetItemGroup(control);

            if (group != null && group.Name != null)
            { 
                string groupName = group.Name;
                RemoveTag(groupName);
            }
        }


        public void OnConfigurationClicked(IRibbonControl control)
        {
            var window = new ConfigWindow(ConfigurationManager.Instance.Configuration, true, this);
            window.ShowDialog();
        }

        public void OnDetachClicked(IRibbonControl control)
        {
            var window = new ConfigWindow(ConfigurationManager.Instance.Configuration, false, this);
            window.Show();
        }


        public void OnOutlookCalendarClicked(IRibbonControl control)
        {
            List<string> names = OutlookCalendarManager.GetAttendees();

            if (names == null || names.Count == 0) return;
            if (ConfigurationManager.Instance.Configuration == null) return;
            if (ConfigurationManager.Instance.Configuration.ItemGroups == null) return;
            if (ConfigurationManager.Instance.Configuration.ItemGroups.Count == 0) return;
            if (ConfigurationManager.Instance.Configuration.ItemGroups[0] == null) return;
            if (ConfigurationManager.Instance.Configuration.ItemGroups[0].Items == null) return;
            if (ConfigurationManager.Instance.Configuration.ItemGroups[0].Items.Count == 0) return;

            int maxNames = Math.Min(names.Count, ConfigurationManager.Instance.Configuration.ItemGroups[0].Items.Count);

            for(int i=0; i< ConfigurationManager.Instance.Configuration.ItemGroups[0].Items.Count; ++i )
            {
                string name = i < names.Count ? names[i] : "Name " + i;
                ConfigurationManager.Instance.Configuration.ItemGroups[0].Items[i].Name = name;
            }

            OnConfigurationChanged(this, null);


        }


        public void OnButtonAction(IRibbonControl control)
        {
            Item item = GetItem(control);
            ItemGroup group = GetItemGroup(control);

            if (group == null || item == null) return;

            SetTag(item, group);
        }


        public void OnEditBoxChanged(IRibbonControl control, string text)
        {
            Item item = GetItem(control);
            item.Name = text;
            ConfigurationManager.Instance.NotifyConfigurationChanged(null);
        }



        public void SetTag(Item item, ItemGroup group)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Slide currentSlide = null;

            if (application.SlideShowWindows != null && application.SlideShowWindows.Count > 0)
            {
                // Presentation mode (slide show)
                currentSlide = application.SlideShowWindows[1].View.Slide;
            } 
            else if (application.ActiveWindow != null && application.ActiveWindow.View != null)
            {
                // Normal mode (edit mode)
                currentSlide = (PowerPoint.Slide)application.ActiveWindow.View.Slide;
            }
            

            if (currentSlide == null)
            {
                return;
            }

            PowerPoint.Shape targetShape = null;

            // Search for a text box with the matching tag
            foreach (PowerPoint.Shape shape in currentSlide.Shapes)
            {
                if (shape.Tags.Count > 0 && shape.Tags["TagID"] == group.Name)
                {
                    targetShape = shape;
                    break;
                }
            }

            // If no text box with the specified tag exists, create a new one
            if (targetShape == null)
            {
                float slideWidth = currentSlide.Master.Width;

                targetShape = currentSlide.Shapes.AddTextbox(
                    Orientation: Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: slideWidth - group.Shape.Right - group.Shape.Width,
                    Top: group.Shape.Top,
                    Width: group.Shape.Width,
                    Height: group.Shape.Hight
                );

                // Tag the text box with the specified id
                targetShape.Tags.Add("TagID", group.Name);
            }

            // Set the text content of the text box
            if (targetShape.TextFrame != null && targetShape.TextFrame.TextRange != null)
            {
                PPT.ReviewTab.Code.Model.ColorScheme colorScheme = PPT.ReviewTab.Code.Model.ColorScheme.Combine(item.ColorScheme, group.Shape.ColorScheme, ConfigurationManager.Instance.Configuration.DefaultColorScheme);
                var textFrame = targetShape.TextFrame;
                textFrame.TextRange.Text = item.Name;

                // Remove bullets
                textFrame.TextRange.ParagraphFormat.Bullet.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Set font size and text color
                textFrame.TextRange.Font.Size = group.Shape.FontSize;
                textFrame.TextRange.Font.Color.RGB = colorScheme.TextColor.RGB();

                // Set margins
                textFrame.MarginLeft = 10; // Left margin
                textFrame.MarginRight = 10; // Right margin
                textFrame.MarginTop = 5; // Top margin
                textFrame.MarginBottom = 5; // Bottom margin


                
                // Set background color using RGB values
                targetShape.Fill.ForeColor.RGB = colorScheme.BackgroundColor.RGB();

                // Set border color using RGB values
                targetShape.Line.ForeColor.RGB = colorScheme.FrameColor.RGB();
                targetShape.Line.Weight = 1; // Border thickness
            }
        }

        


        public void RemoveTag(string id)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Slide currentSlide = (PowerPoint.Slide)application.ActiveWindow.View.Slide;

            if (currentSlide == null)
            {
                return;
            }

            // Search for a shape with the specified tag
            foreach (PowerPoint.Shape shape in currentSlide.Shapes)
            {
                if (shape.Tags.Count > 0 && shape.Tags["TagID"] == id)
                {
                    shape.Delete(); // Remove the shape from the slide
                    return; // Exit the function after removing the shape
                }
            }

            // Optional: Notify if no shape was found
            System.Windows.Forms.MessageBox.Show($"No shape with tag '{id}' was found.", "Shape Not Found");
        }




        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            string rText = GetResourceText("PPT.ReviewTab.ReviewRibbonTab.xml");
            string dynamicContent = CreateDynamicGroups();
            rText = rText.Replace("<DYNAMIC />", dynamicContent);
            return rText;
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.



        /// <summary>
        /// Load the content o th ribbon.
        /// (connected to the onLoad-event of customUI in the ReviewRibbonTab.xml)
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }



        #endregion

        #region Hilfsprogramme

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
