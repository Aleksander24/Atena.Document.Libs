using System;
using System.IO;
using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using System.Collections.Generic;
using System.Linq;
using Syncfusion.Drawing;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Atena.SupportLibs.DocGenerators.ActitvityAnalysis_Word.GroupsData;


namespace Atena.SupportLibs.DocGenerators.ActitvityAnalysis_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";
        public string Label => "DemoTest_ActivityAnalysis";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;


        #region string PROP
        string _head;
        string _headObjects12;
        List<Objects_1_2> _objects12s;
        string _headObjectsVec;
        List<ObjectsVec> _objectsVecs;
        string _headobjectsVis;
        List<ObjectsVis> _objectsViss;
        string _headobjectsLs;
        List<ObjectsLs> _objectsLss;
        string _headobjectsEvpol;
        List<ObjectsEvpol> _objectsEvpols;
        string _headobjectsSamoO;
        List<ObjectsSamoO> _objectsSamoOs;
        string _headobjectsEnPr;
        List<ObjectsEnPr> _objectsEnPrs;
        string _headvehiclesFO;
        List<VehiclesFO> _vehiclesFOs;
        string _headvehiclesPO;
        List<VehiclesPO> _vehiclesPOs;
        string _headvehiclesMunicipalityJP;
        List<VehiclesMunicipalityJP> _vehiclesMunicipalityJPs;
        string _sumAllText;
        #endregion

        #region DocumentGenerator
        public DocumentGenerator(
            string aHead,
            string aHeadObjects12,
            List<Objects_1_2> aObjects12s,
            string aHeadObjectsVec,
            List<ObjectsVec> aObjectsVecs,
            string aHeadObjectsVis,
            List<ObjectsVis> aObjectsViss,
            string aHeadObjectsLs,
            List<ObjectsLs> aObjectsLss,
            string aHeadObjectsEvpol,
            List<ObjectsEvpol> aObjectsEvpols,
            string aHeadObjectsSamoO,
            List<ObjectsSamoO> aObjectsSamoOs,
            string aHeadObjectsEnPr,
            List<ObjectsEnPr> aObjectsEnPrs,
            string aHeadVehiclesFO,
            List<VehiclesFO> aVehiclesFOs,
            string aHeadVehiclesPO,
            List<VehiclesPO> aVehiclesPOs,
            string aHeadVehiclesMunicipality,
            List<VehiclesMunicipalityJP> aVehiclesMunicipalityJPs,
            string aSummAllText)
        {
            _head = aHead;
            _headObjects12 = aHeadObjects12;
            _objects12s = aObjects12s;
            _headObjectsVec = aHeadObjectsVec;
            _objectsVecs = aObjectsVecs;
            _headobjectsVis = aHeadObjectsVis;
            _objectsViss = aObjectsViss;
            _headobjectsLs = aHeadObjectsLs;
            _objectsLss = aObjectsLss;
            _headobjectsEvpol = aHeadObjectsEvpol;
            _objectsEvpols = aObjectsEvpols;
            _headobjectsSamoO = aHeadObjectsSamoO;
            _objectsSamoOs = aObjectsSamoOs;
            _headobjectsEnPr = aHeadObjectsEnPr;
            _objectsEnPrs = aObjectsEnPrs;
            _headvehiclesFO = aHeadVehiclesFO;
            _vehiclesFOs = aVehiclesFOs;
            _headvehiclesPO = aHeadVehiclesPO;
            _vehiclesPOs = aVehiclesPOs;
            _headvehiclesMunicipalityJP = aHeadVehiclesMunicipality;
            _vehiclesMunicipalityJPs = aVehiclesMunicipalityJPs;
            _sumAllText = aSummAllText;
        }
        #endregion


        public byte[] Generate()
        {
            #region Creating document, section, style, paragraph
            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();

            section.PageSetup.Margins.All = 40;

            section.PageSetup.PageSize = new SizeF(575, 792);

            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 11f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.LineSpacing = 10f;
            style.CharacterFormat.TextColor = Color.Black;

            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
            #endregion


            paragraph = SetHead(section);

            SetHeadTable(section);

            #region Align text in first table (right and center alignment)
            WTable table = section.Tables[0] as WTable;
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    foreach (WParagraph paragraph1 in cell.Paragraphs)
                    {
                        if (paragraph1.Text.Contains("UREDBA"))
                            paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                        if (paragraph1.Text.Contains("SPS"))
                            paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                        if (paragraph1.Text.Contains("Število odobrenih"))
                            paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                    }
                }
            }
            #endregion

            SetSecondPartTable(section);

            #region Sums = OBJECTS 1+2
            IWTable table3 = section.AddTable();
            table3.ResetCells(3, 8);
            table3.TableFormat.BackColor = Color.White;
            table3.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table3.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba = _objects12s.Sum(p => p.OdobrenoUredba);
            decimal sumOfAllNakazanoUredba = _objects12s.Sum(p => p.NakazanoUredba);
            decimal sumofAllOdobrenoSPS = _objects12s.Sum(p => p.OdobrenoSPS);
            decimal sumofAllNakazanoSPS = _objects12s.Sum(p => p.NakazanoSPS);
            decimal sumOfAllVlog = _objects12s.Sum(p => p.Vlog);
            decimal sumOfAllNalozb = _objects12s.Sum(p => p.Nalozb);

            table3[0, 0].Width = 80f;
            IWTextRange textRange = table3[0, 0].AddParagraph().AppendText(_headObjects12);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table3[0, 1].Width = 40f;
            table3[0, 2].Width = 70f;
            textRange = table3[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table3[0, 3].Width = 70f;
            textRange = table3[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table3[0, 4].Width = 70f;
            textRange = table3[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table3[0, 5].Width = 70f;
            textRange = table3[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table3[0, 6].Width = 50;
            textRange = table3[0, 6].AddParagraph().AppendText(sumOfAllVlog.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table3[0, 7].Width = 50;
            textRange = table3[0, 7].AddParagraph().AppendText(sumOfAllNalozb.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // Objects 1+2 data
            for (int i = 0; i < _objects12s.Count; i++)
            {
                table3[i + 1, 0].Width = 80f;
                table3[i + 1, 1].Width = 40;
                table3[i + 1, 1].AddParagraph().AppendText(_objects12s[i].Leto.ToString());
                table3[i + 1, 2].Width = 70;
                table3[i + 1, 2].AddParagraph().AppendText(_objects12s[i].OdobrenoUredba.ToString());
                table3[i + 1, 3].Width = 70;
                table3[i + 1, 3].AddParagraph().AppendText(_objects12s[i].NakazanoUredba.ToString());
                table3[i + 1, 4].Width = 70;
                table3[i + 1, 4].AddParagraph().AppendText(_objects12s[i].OdobrenoSPS.ToString());
                table3[i + 1, 5].Width = 70;
                table3[i + 1, 5].AddParagraph().AppendText(_objects12s[i].NakazanoSPS.ToString());
                table3[i + 1, 6].Width = 50;
                table3[i + 1, 6].AddParagraph().AppendText(_objects12s[i].Vlog.ToString());
                table3[i + 1, 7].Width = 50;
                table3[i + 1, 7].AddParagraph().AppendText(_objects12s[i].Nalozb.ToString());
            }

            #region Sum = OBJECTS VEC
            IWTable table4 = section.AddTable();
            table4.ResetCells(3, 8);
            table4.TableFormat.BackColor = Color.White;
            table4.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table4.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba1 = _objectsVecs.Sum(p => p.OdobrenoUredba1);
            decimal sumOfAllNakazanoUredba1 = _objectsVecs.Sum(p => p.NakazanoUredba1);
            decimal sumofAllOdobrenoSPS1 = _objectsVecs.Sum(p => p.OdobrenoSPS1);
            decimal sumofAllNakazanoSPS1 = _objectsVecs.Sum(p => p.NakazanoSPS1);
            decimal sumOfAllVlog1 = _objectsVecs.Sum(p => p.Vlog1);
            decimal sumOfAllNalozb1 = _objectsVecs.Sum(p => p.Nalozb1);

            table4[0, 0].Width = 80f;
            textRange = table4[0, 0].AddParagraph().AppendText(_headObjectsVec);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table4[0, 1].Width = 40f;
            table4[0, 2].Width = 70f;
            textRange = table4[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba1.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table4[0, 3].Width = 70f;
            textRange = table4[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba1.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table4[0, 4].Width = 70f;
            textRange = table4[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS1.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table4[0, 5].Width = 70f;
            textRange = table4[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS1.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table4[0, 6].Width = 50;
            textRange = table4[0, 6].AddParagraph().AppendText(sumOfAllVlog1.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table4[0, 7].Width = 50;
            textRange = table4[0, 7].AddParagraph().AppendText(sumOfAllNalozb1.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // ObjectsVec data
            for (int i = 0; i < _objectsVecs.Count; i++)
            {
                table4[i + 1, 0].Width = 80f;
                table4[i + 1, 1].Width = 40;
                table4[i + 1, 1].AddParagraph().AppendText(_objectsVecs[i].Leto1.ToString());
                table4[i + 1, 2].Width = 70;
                table4[i + 1, 2].AddParagraph().AppendText(_objectsVecs[i].OdobrenoUredba1.ToString());
                table4[i + 1, 3].Width = 70;
                table4[i + 1, 3].AddParagraph().AppendText(_objectsVecs[i].NakazanoUredba1.ToString());
                table4[i + 1, 4].Width = 70;
                table4[i + 1, 4].AddParagraph().AppendText(_objectsVecs[i].OdobrenoSPS1.ToString());
                table4[i + 1, 5].Width = 70;
                table4[i + 1, 5].AddParagraph().AppendText(_objectsVecs[i].NakazanoSPS1.ToString());
                table4[i + 1, 6].Width = 50;
                table4[i + 1, 6].AddParagraph().AppendText(_objectsVecs[i].Vlog1.ToString());
                table4[i + 1, 7].Width = 50;
                table4[i + 1, 7].AddParagraph().AppendText(_objectsVecs[i].Nalozb1.ToString());
            }

            #region Sums = OBJECTS VIS
            IWTable table5 = section.AddTable();
            table5.ResetCells(3, 8);
            table5.TableFormat.BackColor = Color.White;
            table5.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table5.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba2 = _objectsViss.Sum(p => p.OdobrenoUredba2);
            decimal sumOfAllNakazanoUredba2 = _objectsViss.Sum(p => p.NakazanoUredba2);
            decimal sumofAllOdobrenoSPS2 = _objectsViss.Sum(p => p.OdobrenoSPS2);
            decimal sumofAllNakazanoSPS2 = _objectsViss.Sum(p => p.NakazanoSPS2);
            decimal sumOfAllVlog2 = _objectsViss.Sum(p => p.Vlog2);
            decimal sumOfAllNalozb2 = _objectsViss.Sum(p => p.Nalozb2);

            table5[0, 0].Width = 80f;
            textRange = table5[0, 0].AddParagraph().AppendText(_headobjectsVis);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table5[0, 1].Width = 40f;
            table5[0, 2].Width = 70f;
            textRange = table5[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba2.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table5[0, 3].Width = 70f;
            textRange = table5[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba2.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table5[0, 4].Width = 70f;
            textRange = table5[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS2.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table5[0, 5].Width = 70f;
            textRange = table5[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS2.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table5[0, 6].Width = 50;
            textRange = table5[0, 6].AddParagraph().AppendText(sumOfAllVlog2.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table5[0, 7].Width = 50;
            textRange = table5[0, 7].AddParagraph().AppendText(sumOfAllNalozb2.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // ObjectsVis data
            for (int i = 0; i < _objectsViss.Count; i++)
            {
                table5[i + 1, 0].Width = 80f;
                table5[i + 1, 1].Width = 40;
                table5[i + 1, 1].AddParagraph().AppendText(_objectsViss[i].Leto2.ToString());
                table5[i + 1, 2].Width = 70;
                table5[i + 1, 2].AddParagraph().AppendText(_objectsViss[i].OdobrenoUredba2.ToString());
                table5[i + 1, 3].Width = 70;
                table5[i + 1, 3].AddParagraph().AppendText(_objectsViss[i].NakazanoUredba2.ToString());
                table5[i + 1, 4].Width = 70;
                table5[i + 1, 4].AddParagraph().AppendText(_objectsViss[i].OdobrenoSPS2.ToString());
                table5[i + 1, 5].Width = 70;
                table5[i + 1, 5].AddParagraph().AppendText(_objectsViss[i].NakazanoSPS2.ToString());
                table5[i + 1, 6].Width = 50;
                table5[i + 1, 6].AddParagraph().AppendText(_objectsViss[i].Vlog2.ToString());
                table5[i + 1, 7].Width = 50;
                table5[i + 1, 7].AddParagraph().AppendText(_objectsViss[i].Nalozb2.ToString());
            }

            #region Sums = OBJECTS LS
            IWTable table6 = section.AddTable();
            table6.ResetCells(3, 8);
            table6.TableFormat.BackColor = Color.White;
            table6.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table6.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba3 = _objectsLss.Sum(p => p.OdobrenoUredba3);
            decimal sumOfAllNakazanoUredba3 = _objectsLss.Sum(p => p.NakazanoUredba3);
            decimal sumofAllOdobrenoSPS3 = _objectsLss.Sum(p => p.OdobrenoSPS3);
            decimal sumofAllNakazanoSPS3 = _objectsLss.Sum(p => p.NakazanoSPS3);
            decimal sumOfAllVlog3 = _objectsLss.Sum(p => p.Vlog3);
            decimal sumOfAllNalozb3 = _objectsLss.Sum(p => p.Nalozb3);

            table6[0, 0].Width = 80f;
            textRange = table6[0, 0].AddParagraph().AppendText(_headobjectsLs);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table6[0, 1].Width = 40f;
            table6[0, 2].Width = 70f;
            textRange = table6[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba3.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table6[0, 3].Width = 70f;
            textRange = table6[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba3.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table6[0, 4].Width = 70f;
            textRange = table6[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS3.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table6[0, 5].Width = 70f;
            textRange = table6[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS3.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table6[0, 6].Width = 50;
            textRange = table6[0, 6].AddParagraph().AppendText(sumOfAllVlog3.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table6[0, 7].Width = 50;
            textRange = table6[0, 7].AddParagraph().AppendText(sumOfAllNalozb3.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // ObjectsLss data
            for (int i = 0; i < _objectsLss.Count; i++)
            {
                table6[i + 1, 0].Width = 80f;
                table6[i + 1, 1].Width = 40;
                table6[i + 1, 1].AddParagraph().AppendText(_objectsLss[i].Leto3.ToString());
                table6[i + 1, 2].Width = 70;
                table6[i + 1, 2].AddParagraph().AppendText(_objectsLss[i].OdobrenoUredba3.ToString());
                table6[i + 1, 3].Width = 70;
                table6[i + 1, 3].AddParagraph().AppendText(_objectsLss[i].NakazanoUredba3.ToString());
                table6[i + 1, 4].Width = 70;
                table6[i + 1, 4].AddParagraph().AppendText(_objectsLss[i].OdobrenoSPS3.ToString());
                table6[i + 1, 5].Width = 70;
                table6[i + 1, 5].AddParagraph().AppendText(_objectsLss[i].NakazanoSPS3.ToString());
                table6[i + 1, 6].Width = 50;
                table6[i + 1, 6].AddParagraph().AppendText(_objectsLss[i].Vlog3.ToString());
                table6[i + 1, 7].Width = 50;
                table6[i + 1, 7].AddParagraph().AppendText(_objectsLss[i].Nalozb3.ToString());
            }

            #region Sums = OBJECTS EVPOL
            IWTable table7 = section.AddTable();
            table7.ResetCells(3, 8);
            table7.TableFormat.BackColor = Color.White;
            table7.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table7.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba4 = _objectsEvpols.Sum(p => p.OdobrenoUredba4);
            decimal sumOfAllNakazanoUredba4 = _objectsEvpols.Sum(p => p.NakazanoUredba4);
            decimal sumofAllOdobrenoSPS4 = _objectsEvpols.Sum(p => p.OdobrenoSPS4);
            decimal sumofAllNakazanoSPS4 = _objectsEvpols.Sum(p => p.NakazanoSPS4);
            decimal sumOfAllVlog4 = _objectsEvpols.Sum(p => p.Vlog4);
            decimal sumOfAllNalozb4 = _objectsEvpols.Sum(p => p.Nalozb4);

            table7[0, 0].Width = 80f;
            textRange = table7[0, 0].AddParagraph().AppendText(_headobjectsEvpol);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table7[0, 1].Width = 40f;
            table7[0, 2].Width = 70f;
            textRange = table7[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba4.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table7[0, 3].Width = 70f;
            textRange = table7[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba4.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table7[0, 4].Width = 70f;
            textRange = table7[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS4.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table7[0, 5].Width = 70f;
            textRange = table7[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS4.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table7[0, 6].Width = 50;
            textRange = table7[0, 6].AddParagraph().AppendText(sumOfAllVlog4.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table7[0, 7].Width = 50;
            textRange = table7[0, 7].AddParagraph().AppendText(sumOfAllNalozb4.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // ObjectsEvpol data
            for (int i = 0; i < _objectsEvpols.Count; i++)
            {
                table7[i + 1, 0].Width = 80f;
                table7[i + 1, 1].Width = 40;
                table7[i + 1, 1].AddParagraph().AppendText(_objectsEvpols[i].Leto4.ToString());
                table7[i + 1, 2].Width = 70;
                table7[i + 1, 2].AddParagraph().AppendText(_objectsEvpols[i].OdobrenoUredba4.ToString());
                table7[i + 1, 3].Width = 70;
                table7[i + 1, 3].AddParagraph().AppendText(_objectsEvpols[i].NakazanoUredba4.ToString());
                table7[i + 1, 4].Width = 70;
                table7[i + 1, 4].AddParagraph().AppendText(_objectsEvpols[i].OdobrenoSPS4.ToString());
                table7[i + 1, 5].Width = 70;
                table7[i + 1, 5].AddParagraph().AppendText(_objectsEvpols[i].NakazanoSPS4.ToString());
                table7[i + 1, 6].Width = 50;
                table7[i + 1, 6].AddParagraph().AppendText(_objectsEvpols[i].Vlog4.ToString());
                table7[i + 1, 7].Width = 50;
                table7[i + 1, 7].AddParagraph().AppendText(_objectsEvpols[i].Nalozb4.ToString());
            }

            #region Sums = OBJECTS SamoO
            IWTable table8 = section.AddTable();
            table8.ResetCells(3, 8);
            table8.TableFormat.BackColor = Color.White;
            table8.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table8.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba5 = _objectsSamoOs.Sum(p => p.OdobrenoUredba5);
            decimal sumOfAllNakazanoUredba5 = _objectsSamoOs.Sum(p => p.NakazanoUredba5);
            decimal sumofAllOdobrenoSPS5 = _objectsSamoOs.Sum(p => p.OdobrenoSPS5);
            decimal sumofAllNakazanoSPS5 = _objectsSamoOs.Sum(p => p.NakazanoSPS5);
            decimal sumOfAllVlog5 = _objectsSamoOs.Sum(p => p.Vlog5);
            decimal sumOfAllNalozb5 = _objectsSamoOs.Sum(p => p.Nalozb5);

            table8[0, 0].Width = 80f;
            textRange = table8[0, 0].AddParagraph().AppendText(_headobjectsSamoO);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table8[0, 1].Width = 40f;
            table8[0, 2].Width = 70f;
            textRange = table8[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba5.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table8[0, 3].Width = 70f;
            textRange = table8[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba5.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table8[0, 4].Width = 70f;
            textRange = table8[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS5.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table8[0, 5].Width = 70f;
            textRange = table8[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS5.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table8[0, 6].Width = 50;
            textRange = table8[0, 6].AddParagraph().AppendText(sumOfAllVlog5.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table8[0, 7].Width = 50;
            textRange = table8[0, 7].AddParagraph().AppendText(sumOfAllNalozb5.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // ObjectsSamoO data
            for (int i = 0; i < _objectsSamoOs.Count; i++)
            {
                table8[i + 1, 0].Width = 80f;
                table8[i + 1, 1].Width = 40;
                table8[i + 1, 1].AddParagraph().AppendText(_objectsSamoOs[i].Leto5.ToString());
                table8[i + 1, 2].Width = 70;
                table8[i + 1, 2].AddParagraph().AppendText(_objectsSamoOs[i].OdobrenoUredba5.ToString());
                table8[i + 1, 3].Width = 70;
                table8[i + 1, 3].AddParagraph().AppendText(_objectsSamoOs[i].NakazanoUredba5.ToString());
                table8[i + 1, 4].Width = 70;
                table8[i + 1, 4].AddParagraph().AppendText(_objectsSamoOs[i].OdobrenoSPS5.ToString());
                table8[i + 1, 5].Width = 70;
                table8[i + 1, 5].AddParagraph().AppendText(_objectsSamoOs[i].NakazanoSPS5.ToString());
                table8[i + 1, 6].Width = 50;
                table8[i + 1, 6].AddParagraph().AppendText(_objectsSamoOs[i].Vlog5.ToString());
                table8[i + 1, 7].Width = 50;
                table8[i + 1, 7].AddParagraph().AppendText(_objectsSamoOs[i].Nalozb5.ToString());
            }

            #region Sums = OBJECTS EnPr
            IWTable table9 = section.AddTable();
            table9.ResetCells(3, 8);
            table9.TableFormat.BackColor = Color.White;
            table9.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table9.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba6 = _objectsEnPrs.Sum(p => p.OdobrenoUredba6);
            decimal sumOfAllNakazanoUredba6 = _objectsEnPrs.Sum(p => p.NakazanoUredba6);
            decimal sumofAllOdobrenoSPS6 = _objectsEnPrs.Sum(p => p.OdobrenoSPS6);
            decimal sumofAllNakazanoSPS6 = _objectsEnPrs.Sum(p => p.NakazanoSPS6);
            decimal sumOfAllVlog6 = _objectsEnPrs.Sum(p => p.Vlog6);
            decimal sumOfAllNalozb6 = _objectsEnPrs.Sum(p => p.Nalozb6);

            table9[0, 0].Width = 80f;
            textRange = table9[0, 0].AddParagraph().AppendText(_headobjectsEnPr);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table9[0, 1].Width = 40f;
            table9[0, 2].Width = 70f;
            textRange = table9[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba6.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table9[0, 3].Width = 70f;
            textRange = table9[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba6.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table9[0, 4].Width = 70f;
            textRange = table9[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS6.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table9[0, 5].Width = 70f;
            textRange = table9[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS6.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table9[0, 6].Width = 50;
            textRange = table9[0, 6].AddParagraph().AppendText(sumOfAllVlog6.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table9[0, 7].Width = 50;
            textRange = table9[0, 7].AddParagraph().AppendText(sumOfAllNalozb6.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // ObjectsEnPr data
            for (int i = 0; i < _objectsEnPrs.Count; i++)
            {
                table9[i + 1, 0].Width = 80f;
                table9[i + 1, 1].Width = 40;
                table9[i + 1, 1].AddParagraph().AppendText(_objectsEnPrs[i].Leto6.ToString());
                table9[i + 1, 2].Width = 70;
                table9[i + 1, 2].AddParagraph().AppendText(_objectsEnPrs[i].OdobrenoUredba6.ToString());
                table9[i + 1, 3].Width = 70;
                table9[i + 1, 3].AddParagraph().AppendText(_objectsEnPrs[i].NakazanoUredba6.ToString());
                table9[i + 1, 4].Width = 70;
                table9[i + 1, 4].AddParagraph().AppendText(_objectsEnPrs[i].OdobrenoSPS6.ToString());
                table9[i + 1, 5].Width = 70;
                table9[i + 1, 5].AddParagraph().AppendText(_objectsEnPrs[i].NakazanoSPS6.ToString());
                table9[i + 1, 6].Width = 50;
                table9[i + 1, 6].AddParagraph().AppendText(_objectsEnPrs[i].Vlog6.ToString());
                table9[i + 1, 7].Width = 50;
                table9[i + 1, 7].AddParagraph().AppendText(_objectsEnPrs[i].Nalozb6.ToString());
            }

            #region Sums = VEHICLES FO
            IWTable table10 = section.AddTable();
            table10.ResetCells(3, 8);
            table10.TableFormat.BackColor = Color.White;
            table10.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table10.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba7 = _vehiclesFOs.Sum(p => p.OdobrenoUredba7);
            decimal sumOfAllNakazanoUredba7 = _vehiclesFOs.Sum(p => p.NakazanoUredba7);
            decimal sumofAllOdobrenoSPS7 = _vehiclesFOs.Sum(p => p.OdobrenoSPS7);
            decimal sumofAllNakazanoSPS7 = _vehiclesFOs.Sum(p => p.NakazanoSPS7);
            decimal sumOfAllVlog7 = _vehiclesFOs.Sum(p => p.Vlog7);
            decimal sumOfAllNalozb7 = _vehiclesFOs.Sum(p => p.Nalozb7);

            table10[0, 0].Width = 80f;
            textRange = table10[0, 0].AddParagraph().AppendText(_headvehiclesFO);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table10[0, 1].Width = 40f;
            table10[0, 2].Width = 70f;
            textRange = table10[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba7.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table10[0, 3].Width = 70f;
            textRange = table10[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba7.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table10[0, 4].Width = 70f;
            textRange = table10[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS7.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table10[0, 5].Width = 70f;
            textRange = table10[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS7.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table10[0, 6].Width = 50;
            textRange = table10[0, 6].AddParagraph().AppendText(sumOfAllVlog7.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table10[0, 7].Width = 50;
            textRange = table10[0, 7].AddParagraph().AppendText(sumOfAllNalozb7.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // Vehicles FO data
            for (int i = 0; i < _vehiclesFOs.Count; i++)
            {
                table10[i + 1, 0].Width = 80f;
                table10[i + 1, 1].Width = 40;
                table10[i + 1, 1].AddParagraph().AppendText(_vehiclesFOs[i].Leto7.ToString());
                table10[i + 1, 2].Width = 70;
                table10[i + 1, 2].AddParagraph().AppendText(_vehiclesFOs[i].OdobrenoUredba7.ToString());
                table10[i + 1, 3].Width = 70;
                table10[i + 1, 3].AddParagraph().AppendText(_vehiclesFOs[i].NakazanoUredba7.ToString());
                table10[i + 1, 4].Width = 70;
                table10[i + 1, 4].AddParagraph().AppendText(_vehiclesFOs[i].OdobrenoSPS7.ToString());
                table10[i + 1, 5].Width = 70;
                table10[i + 1, 5].AddParagraph().AppendText(_vehiclesFOs[i].NakazanoSPS7.ToString());
                table10[i + 1, 6].Width = 50;
                table10[i + 1, 6].AddParagraph().AppendText(_vehiclesFOs[i].Vlog7.ToString());
                table10[i + 1, 7].Width = 50;
                table10[i + 1, 7].AddParagraph().AppendText(_vehiclesFOs[i].Nalozb7.ToString());
            }

            #region Sums = VEHICLES PO
            IWTable table11 = section.AddTable();
            table11.ResetCells(3, 8);
            table11.TableFormat.BackColor = Color.White;
            table11.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table11.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba8 = _vehiclesPOs.Sum(p => p.OdobrenoUredba8);
            decimal sumOfAllNakazanoUredba8 = _vehiclesPOs.Sum(p => p.NakazanoUredba8);
            decimal sumofAllOdobrenoSPS8 = _vehiclesPOs.Sum(p => p.OdobrenoSPS8);
            decimal sumofAllNakazanoSPS8 = _vehiclesPOs.Sum(p => p.NakazanoSPS8);
            decimal sumOfAllVlog8 = _vehiclesPOs.Sum(p => p.Vlog8);
            decimal sumOfAllNalozb8 = _vehiclesPOs.Sum(p => p.Nalozb8);

            table11[0, 0].Width = 80f;
            textRange = table11[0, 0].AddParagraph().AppendText(_headvehiclesPO);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table11[0, 1].Width = 40f;
            table11[0, 2].Width = 70f;
            textRange = table11[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba8.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table11[0, 3].Width = 70f;
            textRange = table11[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba8.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table11[0, 4].Width = 70f;
            textRange = table11[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS8.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table11[0, 5].Width = 70f;
            textRange = table11[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS8.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table11[0, 6].Width = 50;
            textRange = table11[0, 6].AddParagraph().AppendText(sumOfAllVlog8.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table11[0, 7].Width = 50;
            textRange = table11[0, 7].AddParagraph().AppendText(sumOfAllNalozb8.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // Vehicles PO data
            for (int i = 0; i < _vehiclesPOs.Count; i++)
            {
                table11[i + 1, 0].Width = 80f;
                table11[i + 1, 1].Width = 40;
                table11[i + 1, 1].AddParagraph().AppendText(_vehiclesPOs[i].Leto8.ToString());
                table11[i + 1, 2].Width = 70;
                table11[i + 1, 2].AddParagraph().AppendText(_vehiclesPOs[i].OdobrenoUredba8.ToString());
                table11[i + 1, 3].Width = 70;
                table11[i + 1, 3].AddParagraph().AppendText(_vehiclesPOs[i].NakazanoUredba8.ToString());
                table11[i + 1, 4].Width = 70;
                table11[i + 1, 4].AddParagraph().AppendText(_vehiclesPOs[i].OdobrenoSPS8.ToString());
                table11[i + 1, 5].Width = 70;
                table11[i + 1, 5].AddParagraph().AppendText(_vehiclesPOs[i].NakazanoSPS8.ToString());
                table11[i + 1, 6].Width = 50;
                table11[i + 1, 6].AddParagraph().AppendText(_vehiclesPOs[i].Vlog8.ToString());
                table11[i + 1, 7].Width = 50;
                table11[i + 1, 7].AddParagraph().AppendText(_vehiclesPOs[i].Nalozb8.ToString());
            }

            #region Sums = VEHICLES Municipality JP
            IWTable table12 = section.AddTable();
            table12.ResetCells(3, 8);
            table12.TableFormat.BackColor = Color.White;
            table12.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table12.TableFormat.Paddings.All = 2;
            decimal sumOfAllOdobrenoUredba9 = _vehiclesMunicipalityJPs.Sum(p => p.OdobrenoUredba9);
            decimal sumOfAllNakazanoUredba9 = _vehiclesMunicipalityJPs.Sum(p => p.NakazanoUredba9);
            decimal sumofAllOdobrenoSPS9 = _vehiclesMunicipalityJPs.Sum(p => p.OdobrenoSPS9);
            decimal sumofAllNakazanoSPS9 = _vehiclesMunicipalityJPs.Sum(p => p.NakazanoSPS9);
            decimal sumOfAllVlog9 = _vehiclesMunicipalityJPs.Sum(p => p.Vlog9);
            decimal sumOfAllNalozb9 = _vehiclesMunicipalityJPs.Sum(p => p.Nalozb9);

            table12[0, 0].Width = 80f;
            textRange = table12[0, 0].AddParagraph().AppendText(_headvehiclesMunicipalityJP);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table12[0, 1].Width = 40f;
            table12[0, 2].Width = 70f;
            textRange = table12[0, 2].AddParagraph().AppendText(sumOfAllOdobrenoUredba9.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table12[0, 3].Width = 70f;
            textRange = table12[0, 3].AddParagraph().AppendText(sumOfAllNakazanoUredba9.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table12[0, 4].Width = 70f;
            textRange = table12[0, 4].AddParagraph().AppendText(sumofAllOdobrenoSPS9.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table12[0, 5].Width = 70f;
            textRange = table12[0, 5].AddParagraph().AppendText(sumofAllNakazanoSPS9.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table12[0, 6].Width = 50;
            textRange = table12[0, 6].AddParagraph().AppendText(sumOfAllVlog9.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table12[0, 7].Width = 50;
            textRange = table12[0, 7].AddParagraph().AppendText(sumOfAllNalozb9.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion
            // Vehicles MunicipalityJP data
            for (int i = 0; i < _vehiclesMunicipalityJPs.Count; i++)
            {
                table12[i + 1, 0].Width = 80f;
                table12[i + 1, 1].Width = 40;
                table12[i + 1, 1].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].Leto9.ToString());
                table12[i + 1, 2].Width = 70;
                table12[i + 1, 2].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].OdobrenoUredba9.ToString());
                table12[i + 1, 3].Width = 70;
                table12[i + 1, 3].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].NakazanoUredba9.ToString());
                table12[i + 1, 4].Width = 70;
                table12[i + 1, 4].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].OdobrenoSPS9.ToString());
                table12[i + 1, 5].Width = 70;
                table12[i + 1, 5].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].NakazanoSPS9.ToString());
                table12[i + 1, 6].Width = 50;
                table12[i + 1, 6].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].Vlog9.ToString());
                table12[i + 1, 7].Width = 50;
                table12[i + 1, 7].AddParagraph().AppendText(_vehiclesMunicipalityJPs[i].Nalozb9.ToString());
            }

            #region SUMALL
            IWTable table13 = section.AddTable();
            table13.ResetCells(1, 8);
            table13.TableFormat.BackColor = Color.White;
            table13.TableFormat.Paddings.All = 2;
            table13.TableFormat.HorizontalAlignment = RowAlignment.Left;
            decimal SUMALLOdobrenoUredba =
                _objects12s.Sum(p => p.OdobrenoUredba) +
                _objectsVecs.Sum(p => p.OdobrenoUredba1) +
                _objectsViss.Sum(p => p.OdobrenoUredba2) +
                _objectsLss.Sum(p => p.OdobrenoUredba3) +
                _objectsEvpols.Sum(p => p.OdobrenoUredba4) +
                _objectsSamoOs.Sum(p => p.OdobrenoUredba5) +
                _objectsEnPrs.Sum(p => p.OdobrenoUredba6) +
                _vehiclesFOs.Sum(p => p.OdobrenoUredba7) +
                _vehiclesPOs.Sum(p => p.OdobrenoUredba8) +
                _vehiclesMunicipalityJPs.Sum(p => p.OdobrenoUredba9);
            table13[0, 0].Width = 80f;
            textRange = table13[0, 0].AddParagraph().AppendText(_sumAllText);
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            table13[0, 1].Width = 40f;
            table13[0, 2].Width = 70f;
            textRange = table13[0, 2].AddParagraph().AppendText(SUMALLOdobrenoUredba.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;

            decimal SUMALLNakazanoUredba =
                _objects12s.Sum(p => p.NakazanoUredba) +
                _objectsVecs.Sum(p => p.NakazanoUredba1) +
                _objectsViss.Sum(p => p.NakazanoUredba2) +
                _objectsLss.Sum(p => p.NakazanoUredba3) +
                _objectsEvpols.Sum(p => p.NakazanoUredba4) +
                _objectsSamoOs.Sum(p => p.NakazanoUredba5) +
                _objectsEnPrs.Sum(p => p.NakazanoUredba6) +
                _vehiclesFOs.Sum(p => p.NakazanoUredba7) +
                _vehiclesPOs.Sum(p => p.NakazanoUredba8) +
                _vehiclesMunicipalityJPs.Sum(p => p.NakazanoUredba9);
            table13[0, 3].Width = 70f;
            textRange = table13[0, 3].AddParagraph().AppendText(SUMALLNakazanoUredba.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;

            decimal SUMALLOdobrenoSPS =
                _objects12s.Sum(p => p.OdobrenoSPS) +
                _objectsVecs.Sum(p => p.OdobrenoSPS1) +
                _objectsViss.Sum(p => p.OdobrenoSPS2) +
                _objectsLss.Sum(p => p.OdobrenoSPS3) +
                _objectsEvpols.Sum(p => p.OdobrenoSPS4) +
                _objectsSamoOs.Sum(p => p.OdobrenoSPS5) +
                _objectsEnPrs.Sum(p => p.OdobrenoSPS6) +
                _vehiclesFOs.Sum(p => p.OdobrenoSPS7) +
                _vehiclesPOs.Sum(p => p.OdobrenoSPS8) +
                _vehiclesMunicipalityJPs.Sum(p => p.OdobrenoSPS9);
            table13[0, 4].Width = 70f;
            textRange = table13[0, 4].AddParagraph().AppendText(SUMALLOdobrenoSPS.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;

            decimal SUMALLNakazanoSPS =
                _objects12s.Sum(p => p.NakazanoSPS) +
                _objectsVecs.Sum(p => p.NakazanoSPS1) +
                _objectsViss.Sum(p => p.NakazanoSPS2) +
                _objectsLss.Sum(p => p.NakazanoSPS3) +
                _objectsEvpols.Sum(p => p.NakazanoSPS4) +
                _objectsSamoOs.Sum(p => p.NakazanoSPS5) +
                _objectsEnPrs.Sum(p => p.NakazanoSPS6) +
                _vehiclesFOs.Sum(p => p.NakazanoSPS7) +
                _vehiclesPOs.Sum(p => p.NakazanoSPS8) +
                _vehiclesMunicipalityJPs.Sum(p => p.NakazanoSPS9);
            table13[0, 5].Width = 70f;
            textRange = table13[0, 5].AddParagraph().AppendText(SUMALLNakazanoSPS.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;

            decimal SUMALLVlog =
                _objects12s.Sum(p => p.Vlog) +
                _objectsVecs.Sum(p => p.Vlog1) +
                _objectsViss.Sum(p => p.Vlog2) +
                _objectsLss.Sum(p => p.Vlog3) +
                _objectsEvpols.Sum(p => p.Vlog4) +
                _objectsSamoOs.Sum(p => p.Vlog5) +
                _objectsEnPrs.Sum(p => p.Vlog6) +
                _vehiclesFOs.Sum(p => p.Vlog7) +
                _vehiclesPOs.Sum(p => p.Vlog8) +
                _vehiclesMunicipalityJPs.Sum(p => p.Vlog9);
            table13[0, 6].Width = 50f;
            textRange = table13[0, 6].AddParagraph().AppendText(SUMALLVlog.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;

            decimal SUMALLNalozb =
                _objects12s.Sum(p => p.Nalozb) +
                _objectsVecs.Sum(p => p.Nalozb1) +
                _objectsViss.Sum(p => p.Nalozb2) +
                _objectsLss.Sum(p => p.Nalozb3) +
                _objectsEvpols.Sum(p => p.Nalozb4) +
                _objectsSamoOs.Sum(p => p.Nalozb5) +
                _objectsEnPrs.Sum(p => p.Nalozb6) +
                _vehiclesFOs.Sum(p => p.Nalozb7) +
                _vehiclesPOs.Sum(p => p.Nalozb8) +
                _vehiclesMunicipalityJPs.Sum(p => p.Nalozb9);
            table13[0, 7].Width = 50f;
            textRange = table13[0, 7].AddParagraph().AppendText(SUMALLNalozb.ToString());
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 12f;
            #endregion

            #region Saving document in Word
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Doc);
            stream.Position = 0;
            return stream.ToArray();
            #endregion

        }

        private static void SetSecondPartTable(IWSection section)
        {
            IWTable table2 = section.AddTable();
            table2.ResetCells(1, 8);
            table2.TableFormat.BackColor = Color.White;
            table2.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table2.TableFormat.Paddings.All = 2;
            table2[0, 0].Width = 80f;
            IWTextRange textRangeSecondPartTable = table2[0, 0].AddParagraph().AppendText("Skupina");
            table2[0, 1].Width = 40f;
            textRangeSecondPartTable = table2[0, 1].AddParagraph().AppendText("Leto");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
            table2[0, 2].Width = 70f;
            textRangeSecondPartTable = table2[0, 2].AddParagraph().AppendText("Odobreno");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
            table2[0, 3].Width = 70f;
            textRangeSecondPartTable = table2[0, 3].AddParagraph().AppendText("Nakazano");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
            table2[0, 4].Width = 70f;
            textRangeSecondPartTable = table2[0, 4].AddParagraph().AppendText("Odobreno");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
            table2[0, 5].Width = 70f;
            textRangeSecondPartTable = table2[0, 5].AddParagraph().AppendText("Nakazano");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
            table2[0, 6].Width = 50f;
            textRangeSecondPartTable = table2[0, 6].AddParagraph().AppendText("Vlog");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
            table2[0, 7].Width = 50f;
            textRangeSecondPartTable = table2[0, 7].AddParagraph().AppendText("Naložb");
            textRangeSecondPartTable.CharacterFormat.FontSize = 12f;
        }

        private static void SetHeadTable(IWSection section)
        {
            IWTable table1 = section.AddTable();
            table1.ResetCells(1, 3);
            table1.TableFormat.BackColor = Color.White;
            table1.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table1.TableFormat.Paddings.All = 2;
            table1[0, 0].Width = 260f;
            IWTextRange wTextRangeHead = table1[0, 0].AddParagraph().AppendText("UREDBA");
            wTextRangeHead.CharacterFormat.FontSize = 13f;
            table1[0, 1].Width = 140f;
            wTextRangeHead = table1[0, 1].AddParagraph().AppendText("SPS");
            wTextRangeHead.CharacterFormat.FontSize = 13f;
            table1[0, 2].Width = 100f;
            wTextRangeHead = table1[0, 2].AddParagraph().AppendText("Število odobrenih");
            wTextRangeHead.CharacterFormat.FontSize = 13f;
        }

        private IWParagraph SetHead(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange1 = paragraph.AppendText(_head) as WTextRange;
            textRange1.CharacterFormat.FontSize = 14f;
            textRange1.CharacterFormat.FontName = "Calibri";
            textRange1.CharacterFormat.TextColor = Color.Black;
            textRange1.CharacterFormat.Bold = true;
            return paragraph;
        }
    }
}
