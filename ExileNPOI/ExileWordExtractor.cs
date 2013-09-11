using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using ExileNPOI.Attribute;
using ExileNPOI.Common;
using ExileNPOI.Interface;
using ExileNPOI.Mixins;

namespace ExileNPOI
{
    internal class ExileWordExtractor<T> : IExtractor<T> where T : IExilable
    {
        internal ExileDocumentMeta DocumentMeta { get; set; }
        private readonly Body _body;

        internal ExileWordExtractor(Body body)
        {
            DocumentMeta = Utils.GetTypeMatched<T>();
            _body = body;
        }


        public void FillContent(IList<T> dataList)
        {
            // set page margin
            var sectionProps = new SectionProperties();
            var pageMargin = new PageMargin()
            {
                Top = 500,
                Right = 500,
                Bottom = 500,
                Left = 500,
            };
            sectionProps.Append(pageMargin);
            _body.Append(sectionProps);


            // set title
            var pTitle = new Paragraph(new ParagraphProperties(
                new Justification() {Val = JustificationValues.Center}));
            var run = pTitle.AppendChild(new Run());
            var runProperties = run.AppendChild(new RunProperties());
            var fontSize = new FontSize {Val = "40"};
            var bold = new Bold();
            runProperties.AppendChild(fontSize);
            runProperties.AppendChild(bold);
            // runProperties.AppendChild(id);
            run.AppendChild(new Text(DocumentMeta.TitleText));

            _body.AppendChild(pTitle);

            // todo extract such as SetBorder
            // draw table grid style
            var table = new Table();
            var props = new TableProperties(
                new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    }));
            var tableStyle = new TableStyle {Val = "TableGrid"};
            var tableWidth = new TableWidth {Width = "100%", Type = TableWidthUnitValues.Pct};
            var tableFixedLayout = new TableLayout() {Type = TableLayoutValues.Fixed};
            props.Append(tableStyle, tableWidth, tableFixedLayout);
            table.AppendChild(props);

            // build table header
            var trHeader = new TableRow();
            foreach (var h in DocumentMeta.Headers)
            {
                var tc = new TableCell(new Paragraph(new Run(new Text(h.PropertyDescription))));
                if (h.Width != 0)
                {
                    tc.Append(new TableCellProperties(new TableCellWidth
                    {
                        Type = TableWidthUnitValues.Dxa,
                        Width = h.Width.ToString()
                    }));
                }
                trHeader.Append(tc);
            }
            table.AppendChild(trHeader);


            //build table data
            var dataRowIndex = 0;
            foreach (var d in dataList)
            {
                var trDataRow = new TableRow();
                foreach (var h in DocumentMeta.Headers)
                {
                    string value;
                    if (h.ColumnType == ExileColumnType.AutoIndex)
                    {
                        value = (++dataRowIndex).ToString();
                    }
                    else
                    {
                        value = Utils.GetPropValue(d, h.PropertyName).ToString();
                    }

                    //var type = Utils.GetPropType(d, h.PropertyName);
                    var tc = new TableCell(new Paragraph(new Run(new Text(value))));
                    if (h.Width != 0)
                    {
                        tc.Append(new TableCellProperties(new TableCellWidth
                        {
                            Type = TableWidthUnitValues.Dxa,
                            Width = (h.Width*20).ToString()
                        }));
                    }

                    trDataRow.Append(tc);
                }
                table.AppendChild(trDataRow);
            }

            _body.AppendChild(table);
        }
    }
}
