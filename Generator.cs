using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;


public class ParagraphOptions
{
    public bool ItalicsOn = false;
    public bool BoldOn = false;
    public bool UnderlineOn = false;
    public bool H1On = false;
    public bool H2On = false;
    public bool H3On = false;
    public bool H4On = false;
    public bool H5On = false;

}

public class Generator
{
    public string[] switches = new string[]{"h1", "h2", "h3", "h4", "h5", "p", "br", "i", "b", "u", "/h1", "/h2", "/h3", "/h4", "/h5", "/p", "/i", "/b", "/u"};
    public string[] endparaswitches = new string[]{"h1", "h2", "h3", "h4", "h5", "p", "/h1", "/h2", "/h3", "/h4", "/h5", "/p"};

    public List<Paragraph> GetParagraphsFromText(string text)
    {
        return this.GetParagraphsFromText(text, new ParagraphOptions());
    }

    public List<Paragraph> GetParagraphsFromText(string text, ParagraphOptions options)
    {
        List<Paragraph> paras = new List<Paragraph>();
        List<Run> runs = new List<Run>();
        string currentrun = "";
        int i = 0;
        int j = 0;
        while (i < text.Length)
        {
            bool isswitch = false;
            string checkswitch = "";
            if(text[i]=='<' && i < text.Length-2)
            {
                j = i;
                while (j < text.Length)
                {
                    if(text[j]=='>')
                    {
                        // is this a valid switch?
                        // we end this either way
                        i = j;

                        if(switches.Contains(checkswitch.Substring(1,checkswitch.Length-1).ToLower().Trim()))
                        {
                            isswitch = true;
                        }
                        else
                        {
                            checkswitch += text[j];
                        }
                        break;
                    }
                    checkswitch += text[j];
                    j++;
                }
            }

            if(isswitch)
            {
                // do the switch action, do not add the text to the run
                string switchaction = checkswitch.Substring(1,checkswitch.Length-1).ToLower().Trim();
                runs.Add(this.GetRunFromText(currentrun,options));
                this.SetParagraphAction(options, switchaction);
                if(endparaswitches.Contains(switchaction))
                {
                    var para = new Paragraph();
                    foreach (var run in runs )
                        para.AppendChild(run);
                    paras.Add(para);
                    runs.Clear();
                }
                else
                {
                    // 2024-04-24 SNJW add line breaks
                    if(switchaction == "br")
                        runs[runs.Count-1].AppendChild(new Break());
                }
                currentrun = "";
            }
            else
            {
                
                // just add the text to this run
                if(checkswitch.Length>0)
                {
                    currentrun += checkswitch;
                }
                else
                {
                    currentrun += text[i];
                }
            }
            i++;
        }

        if(( currentrun != "" || paras.Count == 0 ) && runs.Count == 0)
        {
            paras.Add(this.GetParagraphFromText(currentrun,options));
            currentrun = "";
        }

        if(currentrun != "")
            runs.Add(this.GetRunFromText(currentrun,options));

        if (runs.Count > 0)
        {
            var para = new Paragraph();
            foreach (var run in runs )
                para.AppendChild(run);
            paras.Add(para);
        }

        
        return paras;
    }

    public Paragraph GetParagraphFromText(string text, ParagraphOptions options)
    {
        Paragraph para = new Paragraph();
        para.Append(this.GetRunFromText(text, options));
        return para;

    }
    public Run GetRunFromText(string text, ParagraphOptions options)
    {
        Run run = new Run();
        if(options.H1On == true)
        {
            run.Append(new RunProperties(new Bold(), new FontSize(){ Val = "20" }));
        }
        else if(options.H2On == true)
        {
            run.Append(new RunProperties(new Bold(), new FontSize(){ Val = "18" }));
        }
        else if(options.H3On == true)
        {
            run.Append(new RunProperties(new Bold(), new FontSize(){ Val = "16" }));
        }
        else if(options.H4On == true)
        {
            run.Append(new RunProperties(new Bold(), new FontSize(){ Val = "14" }));
        }
        else if(options.H5On == true)
        {
            run.Append(new RunProperties(new Bold(), new FontSize(){ Val = "12" }));
        }

        if(options.ItalicsOn == true)
            run.Append(new RunProperties(new Italic()));
        if(options.BoldOn == true)
            run.Append(new RunProperties(new Bold()));
        if(options.UnderlineOn == true)
            run.Append(new RunProperties(new Underline()));
        run.Append(new Text(text){Space = SpaceProcessingModeValues.Preserve});
        return run;
    }
    public void SetParagraphAction(ParagraphOptions options, string switchaction)
    {
        if(switchaction == "h1")
        {
            options.H1On = true;
            options.H2On = false;
            options.H3On = false;
            options.H4On = false;
            options.H5On = false;
        }
        else if(switchaction == "h2")
        {
            options.H1On = false;
            options.H2On = true;
            options.H3On = false;
            options.H4On = false;
            options.H5On = false;
        }
        else if(switchaction == "h3")
        {
            options.H1On = false;
            options.H2On = false;
            options.H3On = true;
            options.H4On = false;
            options.H5On = false;
        }
        else if(switchaction == "h4")
        {
            options.H1On = false;
            options.H2On = false;
            options.H3On = false;
            options.H4On = true;
            options.H5On = false;
        }
        else if(switchaction == "h5")
        {
            options.H1On = false;
            options.H2On = false;
            options.H3On = false;
            options.H4On = false;
            options.H5On = true;
        }
        else if(switchaction == "/h1")
        {
            options.H1On = false;
        }
        else if(switchaction == "/h2")
        {
            options.H2On = false;
        }
        else if(switchaction == "/h3")
        {
            options.H3On = false;
        }
        else if(switchaction == "/h4")
        {
            options.H4On = false;
        }
        else if(switchaction == "/h5")
        {
            options.H5On = false;
        }
        else if(switchaction == "/i")
        {
            options.ItalicsOn = false;
        }
        else if(switchaction == "/b")
        {
            options.BoldOn = false;
        }
        else if(switchaction == "/u")
        {
            options.UnderlineOn = false;
        }
        else if(switchaction == "i")
        {
            options.ItalicsOn = true;
        }
        else if(switchaction == "b")
        {
            options.BoldOn = true;
        }
        else if(switchaction == "u")
        {
            options.UnderlineOn = true;
        }

    }

   

 
    public void CreateDocument(string fileName, List<Plan> lessonplans)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // Create a new section and set page properties
            SectionProperties sectionProps = body.AppendChild(new SectionProperties());
            PageSize pageSize = new PageSize() { Width = (UInt32Value)842 * (UInt32)20, Height = (UInt32Value)595 * (UInt32)20, Orient = PageOrientationValues.Landscape };
            PageMargin pageMargin = new PageMargin() { Top = 720, Right = (UInt32)720, Bottom = 720, Left = (UInt32)720, Header = (UInt32)720, Footer = (UInt32)720, Gutter = (UInt32)0 };

            sectionProps.Append(pageSize);
            sectionProps.Append(pageMargin);

            // Create a new run
            Run run = new Run();

            // Create a new run property
            RunProperties runProps = new RunProperties();

            // Create and set the font size and font type
            RunFonts runFont = new RunFonts();  
            runFont.Ascii = "Calibri";
            FontSize fontSize = new FontSize() { Val = "10" };

            // Append font size and type to run property
            runProps.Append(runFont);
            runProps.Append(fontSize);

            // Append run property to run
            run.AppendChild(runProps);

            // Append run to paragraph, paragraph to body and body to document
            Paragraph paragraph = body.AppendChild(new Paragraph());
            paragraph.AppendChild(run);
            mainPart.Document.Save();

            CreateNumberingPart(mainPart);



            // TO DO
            bool doTlaPlan = false;
            if(doTlaPlan)
            {
                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());
                foreach (var item in TlaPlanHeader(lessonplans))
                    body.AppendChild(item);

                Table tlaPlan = TlaPlan(lessonplans);
                body.AppendChild(tlaPlan);
                //mainPart.Document.Save();

            }


            // var paras =  GetParagraphsFromText("a<b>c",new ParagraphOptions());
            // foreach(var para in paras)
            //     body.AppendChild(para);
            // lessons.Clear();

            body.AppendChild(new Paragraph());
            body.AppendChild(new Paragraph());
            Table unitPlan = UnitPlan(lessonplans);
            body.AppendChild(unitPlan);
            //mainPart.Document.Save();


            foreach(var lesson in lessonplans)
            {
                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());
                Table lessonHeader = LessonHeader(lesson);
                body.AppendChild(lessonHeader);

                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());
                Table lessonContext = LessonContext(lesson);
                body.AppendChild(lessonContext);

                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());
                Table lessonLearning = LessonLearning(lesson);
                body.AppendChild(lessonLearning);

                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());
                Table lessonProcedure = LessonProcedure(lesson);
                body.AppendChild(lessonProcedure);

                body.AppendChild(new Paragraph());
                body.AppendChild(new Paragraph());
                Table lessonEvaluation = LessonEvaluation(lesson);
                body.AppendChild(lessonEvaluation);


            }

            mainPart.Document.Save();
            wordDoc.Dispose();
        }
    }

    private List<Paragraph> TlaPlanHeader(List<Plan> lessonplans)
    {
        List<Paragraph> output = new();
        return output;
//        throw new NotImplementedException();
    }

    public Table TlaPlan(List<Plan> lessonplans)
    {
        Table table = new Table();
        TableProperties props = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableLayout(){ Type = TableLayoutValues.Fixed }
        );
        table.AppendChild(props);


        // TableLayout tl = new TableLayout(){ Type = TableLayoutValues.Fixed };
        // props.TableLayout = tl;



        //         // Add 3 columns to the table.
        //         TableGrid tg = new TableGrid(new GridColumn(new Width()), new GridColumn(), new GridColumn(), new GridColumn(), new GridColumn());
        //         tabletbl.AppendChild(tg);

        // Row 1
        table.Append(new TableRow(
            this.GetHeaderCell("Unit Plan"),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" "))))
        ));

        // Row 2
        table.Append(new TableRow(
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Code"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Description"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Elaboration"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Objectives"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Resources")))))
        ));

        // Row 2
        foreach(var lesson in lessonplans)
        {

            var codeCell = new TableCell(
                this.GetBorderCellProperties(),
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "700" }));
            var descCell = new TableCell(
                this.GetBorderCellProperties(),
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1100" }));
            var elaborationCell = new  TableCell(
                this.GetBorderCellProperties(), 
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1100" }));
            var learningOutcomeCell = new  TableCell(
                this.GetBorderCellProperties(), 
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1200" }));
            var resourcesCell = new  TableCell(
                this.GetBorderCellProperties(), 
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "900" }));

            // if empty or have no values, insert a single space
            if((lesson.LessonCode??"")=="")
            {
                codeCell.Append(new Paragraph(new Run(new Bold(), new Text(" "))));
            }
            else
            {
                codeCell.Append(new Paragraph(new Run(new Bold(), new Text(lesson.LessonCode))));
            }

            if((lesson.KlaTopic??"")=="")
            {
                descCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {

                string topicText = lesson.KlaTopic;
                if((lesson.SchoolCode??"") != "")
                    topicText = topicText + " (" + lesson.SchoolCode + ')';
                descCell.Append(new Paragraph(new Run(new Text(topicText))));
            }

            if(lesson.Elaborations.Count(x => (x.Name??"") != "") == 0)
            {
                elaborationCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                foreach(var elaboration in lesson.Elaborations.Where(x => (x.Name??"") != ""))
                    foreach(var para in this.GetParagraphsFromText(elaboration.Name))
                        elaborationCell.Append(para);
            }

            if(lesson.LearningOutcomes.Count(x => (x.Name??"") != "") == 0)
            {
                learningOutcomeCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                learningOutcomeCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text($"By the end of this lesson, students will know how/be able to:"))));
                foreach(var outcome in lesson.LearningOutcomes.Where(x => (x.Name??"") != ""))
                    foreach(var para in this.GetParagraphsFromText(outcome.Name))
                        learningOutcomeCell.Append(para);
            }

            if(lesson.Resources.Count(x => (x.Name??"") != "") == 0)
            {
                resourcesCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                // resourcesCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text($"By the end of this lesson, students will know how/be able to:"))));
                foreach(var resource in lesson.Resources.Where(x => (x.Name??"") != ""))
                    foreach(var para in this.GetParagraphsFromText(resource.Name))
                        resourcesCell.Append(para);
            }

            table.Append(new TableRow(codeCell, descCell, elaborationCell,learningOutcomeCell,resourcesCell));


        }

        return table;
    }


    public Table UnitPlan(List<Plan> lessonplans)
    {
        Table table = new Table();
        TableProperties props = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableLayout(){ Type = TableLayoutValues.Fixed }
        );
        table.AppendChild(props);


        // TableLayout tl = new TableLayout(){ Type = TableLayoutValues.Fixed };
        // props.TableLayout = tl;



        //         // Add 3 columns to the table.
        //         TableGrid tg = new TableGrid(new GridColumn(new Width()), new GridColumn(), new GridColumn(), new GridColumn(), new GridColumn());
        //         tabletbl.AppendChild(tg);

        // Row 1
        table.Append(new TableRow(
            this.GetHeaderCell("Unit Plan"),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" "))))
        ));

        // Row 2
        table.Append(new TableRow(
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Code"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Description"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Elaboration"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Objectives"))))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Resources")))))
        ));

        // Row 2
        foreach(var lesson in lessonplans)
        {

            var codeCell = new TableCell(
                this.GetBorderCellProperties(),
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "700" }));
            var descCell = new TableCell(
                this.GetBorderCellProperties(),
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1100" }));
            var elaborationCell = new  TableCell(
                this.GetBorderCellProperties(), 
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1100" }));
            var learningOutcomeCell = new  TableCell(
                this.GetBorderCellProperties(), 
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "1200" }));
            var resourcesCell = new  TableCell(
                this.GetBorderCellProperties(), 
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "900" }));

            // if empty or have no values, insert a single space
            if((lesson.LessonCode??"")=="")
            {
                codeCell.Append(new Paragraph(new Run(new Bold(), new Text(" "))));
            }
            else
            {
                codeCell.Append(new Paragraph(new Run(new Bold(), new Text(lesson.LessonCode))));
            }

            if((lesson.KlaTopic??"")=="")
            {
                descCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                string topicText = lesson.KlaTopic;
                if((lesson.SchoolCode??"") != "")
                    topicText = topicText + " (" + lesson.SchoolCode + ')';

                descCell.Append(new Paragraph(new Run(new Text(topicText))));
            }

            if(lesson.Elaborations.Count(x => (x.Name??"") != "") == 0)
            {
                elaborationCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                foreach(var elaboration in lesson.Elaborations.Where(x => (x.Name??"") != ""))
                    foreach(var para in this.GetParagraphsFromText(elaboration.Name))
                        elaborationCell.Append(para);
            }

            if(lesson.LearningOutcomes.Count(x => (x.Name??"") != "") == 0)
            {
                learningOutcomeCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                learningOutcomeCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text($"By the end of this lesson, students will know how/be able to:"))));
                foreach(var outcome in lesson.LearningOutcomes.Where(x => (x.Name??"") != ""))
                    foreach(var para in this.GetParagraphsFromText(outcome.Name))
                        learningOutcomeCell.Append(para);
            }

            if(lesson.Resources.Count(x => (x.Name??"") != "") == 0)
            {
                resourcesCell.Append(new Paragraph(new Run(new Text(" "))));
            }
            else
            {
                // resourcesCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text($"By the end of this lesson, students will know how/be able to:"))));
                foreach(var resource in lesson.Resources.Where(x => (x.Name??"") != ""))
                    foreach(var para in this.GetParagraphsFromText(resource.Name))
                        resourcesCell.Append(para);
            }

            table.Append(new TableRow(codeCell, descCell, elaborationCell,learningOutcomeCell,resourcesCell));


        }

        return table;
    }

    private Table LessonHeader(Plan lesson)
    {
        Table table = new Table();

        TableRow firstRow = new TableRow();

        TableCell firstRowCell1 = new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "9000" },
                new Shading { Fill = "000000", Val = ShadingPatternValues.Clear }),
            new Paragraph(
                new Run(
                    new RunProperties(new Bold(), new Color { Val = "FFFFFF" }),
                    new Text("LESSON PLAN"))));


        firstRowCell1.Append(new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Restart }));
        firstRowCell1.Append(this.GetBorderCellProperties()); 


        TableCellProperties tableCellProperties1 = new TableCellProperties();
        HorizontalMerge verticalMerge1 = new HorizontalMerge()
        {
            Val = MergedCellValues.Continue
        };
        tableCellProperties1.Append(verticalMerge1);

        TableCell firstRowCell2 = new TableCell();
        firstRowCell2.Append(new Paragraph(new Run(new Text("Dummy Cell 1"))));
        firstRowCell2.Append(new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }));
        TableCell firstRowCell3 = new TableCell();
        firstRowCell3.Append(new Paragraph(new Run(new Text("Dummy Cell 2"))));
        firstRowCell3.Append(new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }));
        firstRow.Append(firstRowCell1);
        firstRow.Append(firstRowCell2);
        firstRow.Append(firstRowCell3);

        TableRow secondRow = new TableRow();
        
        TableCell secondRowCell1 = new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" }),
            new Paragraph(new Run(new Text($"Year level: {lesson.YearLevel}"))));
        secondRowCell1.Append(this.GetBorderCellProperties()); 

        secondRow.Append(secondRowCell1);

        TableCell secondRowCell2 = new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" }),
            new Paragraph(new Run(new Text($"Duration: {lesson.Duration}"))));
        secondRowCell2.Append(this.GetBorderCellProperties()); 
        secondRow.Append(secondRowCell2);

        TableCell secondRowCell3 = new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "5000" }));
        
        if((lesson.LessonCode??"") != "")
        {
            secondRowCell3.Append(new Paragraph(new Run(new Text($"Lesson Code: {lesson.LessonCode}"))));
        }
        if((lesson.KlaTopic??"") != "")
        {
            secondRowCell3.Append(new Paragraph(new Run(new Text($"KLA/Topic: "))));

            string topicText = lesson.KlaTopic;
            if((lesson.SchoolCode??"") != "")
                topicText = topicText + " (" + lesson.SchoolCode + ')';
            secondRowCell3.Append(new Paragraph(new Run(new Text($"{topicText}"))));
        }
            

        secondRowCell3.Append(this.GetBorderCellProperties()); 
        secondRow.Append(secondRowCell3);


        table.Append(firstRow);
        table.Append(secondRow);


        // table.Append(secondRow);

        return table;
    }

    public Table LessonContext(Plan lesson)
    {
        Table table = new Table();
        TableProperties props = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.AppendChild(props);

        // Row 1
        table.Append(new TableRow(
            this.GetHeaderCell("Lesson Context"),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue }), new Paragraph(new Run(new Text($" "))))
        ));



        var curriculumCell =  new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Bold(), new Text($"Curriculum:"))));

        var learnersCell = new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Bold(), new Text($"Learners:"))));

        var resourcesCell = new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Bold(), new Text($"Resources:"))));



        // // Row 2
        // var curriculumCell = new TableCell(
        //     this.GetBorderCellProperties(),
        //      new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Restart }),
        //     new Paragraph(new Run(new Bold()), new Text($"Curriculum:")));

        if((lesson.Curriculum??"")!="")
            curriculumCell.Append(new Paragraph(new Run(new Text(lesson.Curriculum))));

        // var learnersCell = new TableCell(
        //     this.GetBorderCellProperties(),
        //     new Paragraph(new Run(new Bold()), new Text($"Learners:")));

        foreach(var learner in lesson.Learners.Where(x => (x.Name??"") != ""))
        {
            string learnertext = learner.Name??"";
            string differentiation = learner.Differentiation??"";
            if(learnertext!= "")
            {
                if(differentiation!= "")
                    learnertext = learnertext + " (" + differentiation + ")";
                foreach(var para in this.GetParagraphsFromText(learnertext))
                    learnersCell.Append(para);
            }
        }

        this.SetFontSizeInTableCell(learnersCell,10);

        // var resourcesCell = new TableCell(
        //     this.GetBorderCellProperties(),
        //     new Paragraph(new Run(new Bold()), new Text($"Resources:")));

        foreach(var resource in lesson.Resources.Where(x => (x.Name??"") != ""))
            foreach(var para in this.GetParagraphsFromText(resource.Name))
                resourcesCell.Append(para);


        this.SetFontSizeInTableCell(resourcesCell,9);

        // Row 4
        table.Append(new TableRow(curriculumCell,learnersCell,resourcesCell));



        // table.Append(new TableRow(curriculumCell,learnersCell,resourcesCell));

        if((lesson?.PedagogicalStrategy??"")!="")
        {
            // Row 3
            table.Append(new TableRow(
                this.GetHeaderCell("Pedagogy"),
                new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" ")))),
                new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" "))))
            ));

            // Row 4
            table.Append(new TableRow(
                new TableCell(this.GetBorderCellProperties(),  new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Restart }), new Paragraph(new Run(new Text($"Main pedagogical strategy: {lesson.PedagogicalStrategy??""}")))),
                new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" ")))),
                new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" "))))
            ));
        }



        return table;
    }
    private Table LessonLearning(Plan lesson)
    {
        Table table = new Table();
        TableProperties props = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.AppendChild(props);

        // Row 1
        table.Append(new TableRow(
            this.GetHeaderCell("Intended & Assessed learning"),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" "))))
        ));


        // Row 2
        var contentDescriptionCell = new  TableCell(this.GetBorderCellProperties(), new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Restart }));
        contentDescriptionCell.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Context description:"))));
        if((lesson.ContentDescription??"")!="")
            contentDescriptionCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text(lesson.ContentDescription))));
        
        if(lesson.Elaborations.Count(x => (x.Name??"") != "")>0)
            contentDescriptionCell.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Relevant elaborations:"))));
        foreach(var elaboration in lesson.Elaborations.Where(x => (x.Name??"") != ""))
            foreach(var para in this.GetParagraphsFromText(elaboration.Name))
                contentDescriptionCell.Append(para);

        table.Append(new TableRow(contentDescriptionCell, new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" "))))));



        // foreach (var item in items)
        // {
        //     var para = new Paragraph(
        //         new ParagraphProperties(
        //             new ParagraphStyleId() { Val = "BodyText" },
        //             new NumberingProperties(
        //                 new NumberingLevelReference() { Val = 0 },
        //                 new NumberingId() { Val = 1 }
        //             )
        //         ),
        //         new Run(new Text(item))
        //     );
        //     mainPart.Document.Body.Append(para);
        // }



// By the end of this lesson, students will know/do):
// - students can describe and model the relative positions of the earth/moon and sun
// - students can describe orbit of earth and moon and sun in relation to each other and the effect this
// has
// - students understand and explain what causes day and night on planets
// Assessment of Learning Outcome (How I will know which students have met the learning outcome?)



        // Row 3
        var learningOutcomeCell = new  TableCell(
            this.GetBorderCellProperties(),
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4000" }),
            new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Learning outcome (learning intent):"))));
        learningOutcomeCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text($"By the end of this lesson, students will know how/be able to:"))));
        foreach(var outcome in lesson.LearningOutcomes.Where(x => (x.Name??"") != ""))
            foreach(var para in this.GetParagraphsFromText(outcome.Name))
                learningOutcomeCell.Append(para);

        var successCriteriaCell = new  TableCell(
            this.GetBorderCellProperties(),
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4000" }),
            new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Evidence of learning when achieved (success criteria):"))));
        successCriteriaCell.Append(new Paragraph(new Run(new RunProperties(new Italic()), new Text($"To succeed in this lesson, a student must be able to:"))));
        foreach(var success in lesson.SuccessCriteria.Where(x => (x.Name??"") != ""))
            foreach(var para in this.GetParagraphsFromText(success.Name))
                successCriteriaCell.Append(para);

        table.Append(new TableRow(learningOutcomeCell, successCriteriaCell));
        return table;
    }
    private Table LessonProcedure(Plan lesson)
    {
        Table table = new Table();
        TableProperties props = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.AppendChild(props);

        // Row 1
        table.Append(new TableRow(
            this.GetHeaderCell("Enacted Learning (Procedure)"),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" "))))
        ));

        // Row 2
        table.Append(new TableRow(
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Time")))),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Stage")))),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Differentiation")))),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new RunProperties(new Bold()), new Text($"Management"))))
        ));

        
        foreach(var stage in lesson.Steps)
        {
            var timeCell = new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"{stage.Time}"))));
            var stageCell = new TableCell(this.GetBorderCellProperties());
            var differentiationCell = new TableCell(this.GetBorderCellProperties());
            var managementCell = new TableCell(this.GetBorderCellProperties());
            foreach(var para in this.GetParagraphsFromText(stage.Stage))
                stageCell.Append(para);
            foreach(var para in this.GetParagraphsFromText(stage.Differentiation))
                differentiationCell.Append(para);
            foreach(var para in this.GetParagraphsFromText(stage.Management))
                managementCell.Append(para);

            table.Append(new TableRow(
                timeCell,
                stageCell,
                differentiationCell,
                managementCell));
        }

        return table;
    }
    private Table LessonEvaluation(Plan lesson)
    {
        Table table = new Table();
        TableProperties props = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.AppendChild(props);

        // Row 1
        table.Append(new TableRow(
            this.GetHeaderCell("Lesson evaluation"),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" ")))),
            new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($" "))))
        ));


        // Row 2
        var successMetCell = new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"Explain the extent students met the success criteria:"))));
        successMetCell.Append(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"{lesson.LessonEvaluation}"))));
        var nextStepsCell = new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"Next steps for learning"))));
        nextStepsCell.Append(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"{lesson.NextSteps}"))));
        var reflectionCell = new TableCell(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"Put teacher reflection here:"))));
        reflectionCell.Append(this.GetBorderCellProperties(), new Paragraph(new Run(new Text($"{lesson.Reflection}"))));
        table.Append(new TableRow(successMetCell, nextStepsCell,reflectionCell));


        return table;
    }




    private void CreateNumberingPart(MainDocumentPart mainPart)
    {
        NumberingDefinitionsPart numberingPart;
        if (mainPart.NumberingDefinitionsPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                new AbstractNum(
                    new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                    new Level(
                        new StartNumberingValue() { Val = 1 },
                        new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        new LevelText() { Val = "%1." },
                        new LevelJustification() { Val = LevelJustificationValues.Left }
                    ),
                    new NumberingSymbolRunProperties(
                        new RunFonts() { ComplexScript = "Times New Roman" }
                    )
                )
                { AbstractNumberId = 0 },
                new NumberingInstance(
                    new AbstractNumId() { Val = 0 }
                )
                { NumberID = 1 }
            );
        }
        else
        {
            numberingPart = mainPart.NumberingDefinitionsPart;
        }



    }

    private TableCellProperties GetBorderCellProperties()
    {
        return new TableCellProperties(
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Color = "000000" }
                ),
                new TableCellProperties(new HorizontalMerge { Val = MergedCellValues.Continue })
            );
    }

    private TableCell GetHeaderCell(string text)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "9000" },
                new Shading { Fill = "A6A6A6", Val = ShadingPatternValues.Clear },
                this.GetBorderCellProperties(),
                new HorizontalMerge { Val = MergedCellValues.Restart }
                ),
            new Paragraph(
                new Run(
                    new RunProperties(new Bold(), new Color { Val = "000000" }),
                    new Text(text))
            )
        );
    }


    // Method to set the font size of paragraphs in a TableCell
    private void SetFontSizeInTableCell(TableCell learnerCell, int fontSize)
    {
        // Define the font size in half-point units (10 point size = 20 half-point units)
        string fontSizeValue = (fontSize * 2).ToString();

        // Create a new RunProperties object to hold the font size property
        RunProperties runProperties = new RunProperties(new FontSize() { Val = fontSizeValue });

        // Apply this RunProperties to each Paragraph in the TableCell
        foreach (var paragraph in learnerCell.Elements<Paragraph>())
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                run.PrependChild(runProperties.CloneNode(true));
            }
        }

        // Create a ParagraphProperties object with the RunProperties to apply to new paragraphs
        ParagraphProperties paragraphProperties = new ParagraphProperties(new ParagraphMarkRunProperties(runProperties.CloneNode(true)));

        // Append this ParagraphProperties to the TableCell
        foreach (var paragraph in learnerCell.Elements<Paragraph>())
        {
            paragraph.ParagraphProperties = paragraphProperties.CloneNode(true) as ParagraphProperties;
        }
    }


}