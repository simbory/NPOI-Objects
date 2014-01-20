using NPOI.Objects.Attributes;
using NPOI.SS.UserModel;

namespace NPOI.Example
{
    [NPOIObject]
    public class TestModel
    {
        [Column(0)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "gray", TextColor = "White")]
        [CellStyle(TextColor = "#eeeeee", BackgroundColor = "#333333")]
        [AlternateCellStyle(TextColor = "Red", BackgroundColor = "Blue")]
        public uint Id { get; set; }

        [Column("Name")]
        [DrawingIgnore]
        [RichText]
        public string NameHtml { get; set; }

        [Column(1)]
        [HeaderStyle(ColumnWidth = 23, BackgroundColor = "#FF0000", TextColor = "#FFFFFF")]
        public string Name { get; set; }

        [Column(2)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "#00FF00", TextColor = "#FFFFFF")]
        [CellStyle(FontFamily = "Matura MT Script Capitals")]
        [AlternateCellStyle(FontWeight = 700)]
        public string From { get; set; }

        [Column(3)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "#0000FF", TextColor = "#FFFFFF")]
        [AlternateCellStyle(FontWeight = 700)]
        public string Type { get; set; }

        [Column(4)]
        [HeaderStyle(ColumnWidth = 40, BackgroundColor = "#000000", TextColor = "#FFFFFF", TextAlign = HorizontalAlignment.Center)]
        public string CityCountry { get; set; }
    }
}
