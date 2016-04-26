using NPOI.Objects;
using NPOI.SS.UserModel;

namespace NPOI.Example
{
    [NPOIObject]
    public class TestModel
    {
        [NPOIColumn(0)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "gray", TextColor = "White")]
        [CellStyle(TextColor = "#eeeeee", BackgroundColor = "#333333")]
        [AlternateCellStyle(TextColor = "Red", BackgroundColor = "Blue")]
        public uint Id { get; set; }

        [NPOIColumn("Name")]
        [DrawingIgnore]
        [RichText]
        public string NameHtml { get; set; }

        [NPOIColumn(1)]
        [HeaderStyle(ColumnWidth = 23, BackgroundColor = "#FF0000", TextColor = "#FFFFFF")]
        public string Name { get; set; }

        [NPOIColumn(2)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "#00FF00", TextColor = "red")]
        [CellStyle(FontFamily = "Matura MT Script Capitals")]
        [AlternateCellStyle(FontWeight = 700)]
        public string From { get; set; }

        [NPOIColumn(3)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "#0000FF", TextColor = "#FFFFFF")]
        [AlternateCellStyle(FontWeight = 700)]
        public string Type { get; set; }

        [NPOIColumn(4)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "#000000", TextColor = "#FFFFFF", TextAlign = HorizontalAlignment.Center)]
        public string CityCountry { get; set; }
    }
}
