using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointProofingLanguageVSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Type langType = typeof(MsoLanguageID);
            int begin = "msoLanguageID".Length;

            MsoLanguageID defaultLanguageID = Globals.ThisAddIn.Application.ActivePresentation.DefaultLanguageID;
            RibbonDropDownItem defaultItem = null;

            Dictionary<MsoLanguageID, string> dict = new Dictionary<MsoLanguageID, string>();
            foreach (MsoLanguageID lang in Enum.GetValues(langType))
            {
                try
                {
                    dict.Add(lang, Enum.GetName(langType, lang).Substring(begin));
                } catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
            dict = dict.OrderBy(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
            foreach(MsoLanguageID lang in dict.Keys)
            {
                var item = this.Factory.CreateRibbonDropDownItem();
                item.Tag = lang;
                item.Label = dict[lang];
                if (lang == defaultLanguageID)
                {
                    defaultItem = item;
                }
                dropDown1.Items.Add(item);
            }

            dropDown1.SelectedItem = defaultItem;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MsoLanguageID lang = (MsoLanguageID) dropDown1.SelectedItem.Tag;
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    ChangeAllSubShapes(shape, lang);
                }
            }

            foreach (CustomLayout cl in presentation.SlideMaster.CustomLayouts)
            {
                foreach (Shape shape in cl.Shapes)
                {
                    ChangeAllSubShapes(shape, lang);
                }
            }

            presentation.DefaultLanguageID = lang;
        }

        private void ChangeAllSubShapes(Shape shape, MsoLanguageID lang)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.LanguageID = lang;
            }

            if (shape.HasTable == MsoTriState.msoTrue)
            {
                for (int i = 1; i <= shape.Table.Rows.Count; i++)
                {
                    for (int j = 1; j <= shape.Table.Columns.Count; j++)
                    {
                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.LanguageID = lang;
                    }
                }
            }

            if (shape.Type == MsoShapeType.msoGroup || shape.Type == MsoShapeType.msoSmartArt)
            {
                foreach (Shape s in shape.GroupItems)
                {
                    ChangeAllSubShapes(s, lang);
                }
            }
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
