using System;
using System.Collections.Generic;

namespace CMRPrint
{
    public class CmrField
    {
        public int PlaceNumber { get; set; }
        public string FieldName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public float DefaultX { get; set; }
        public float DefaultY { get; set; }
        public int DefaultFontSize { get; set; } = 10;
    }

    public static class CmrPlaces
    {
        public static List<CmrField> GetStandardPlaces()
        {
            return new List<CmrField>
            {
                new() { PlaceNumber = 1, FieldName = "ConsignorName", Description = "1. Sender", DefaultX = 40f, DefaultY = 80f, DefaultFontSize = 9 },
                new() { PlaceNumber = 2, FieldName = "ConsignorDetails", Description = "2. Destination", DefaultX = 40f, DefaultY = 130f, DefaultFontSize = 8 },
                new() { PlaceNumber = 3, FieldName = "LoadingInstructions", Description = "3. Place of Delivery Good", DefaultX = 40f, DefaultY = 160f, DefaultFontSize = 8 },
                new() { PlaceNumber = 4, FieldName = "ConsignorRemarks", Description = "4. Place and Date of Reception", DefaultX = 40f, DefaultY = 200f, DefaultFontSize = 8 },
                new() { PlaceNumber = 5, FieldName = "DocumentsAttached", Description = "5. Documents attached", DefaultX = 40f, DefaultY = 240f, DefaultFontSize = 8 },
                new() { PlaceNumber = 6, FieldName = "Seals", Description = "6. Marks and Numbers", DefaultX = 120f, DefaultY = 240f, DefaultFontSize = 8 },
                new() { PlaceNumber = 7, FieldName = "PackagingType", Description = "7. Number of Packages", DefaultX = 200f, DefaultY = 240f, DefaultFontSize = 8 },
                new() { PlaceNumber = 8, FieldName = "GoodsDescription", Description = "8. Goods description", DefaultX = 40f, DefaultY = 280f, DefaultFontSize = 9 },
                new() { PlaceNumber = 9, FieldName = "NatureofGoods", Description = "9. Nature of Goods", DefaultX = 40f, DefaultY = 360f, DefaultFontSize = 9 },
                new() { PlaceNumber = 10, FieldName = "LoadingOrderNumber", Description = "10. Statistical Number", DefaultX = 200f, DefaultY = 360f, DefaultFontSize = 8 },
                new() { PlaceNumber = 11, FieldName = "TransportChargesPlace", Description = "11. Gross Weight", DefaultX = 400f, DefaultY = 80f, DefaultFontSize = 8 },
                new() { PlaceNumber = 12, FieldName = "ConsigeeName", Description = "12. Volume in m3", DefaultX = 400f, DefaultY = 110f, DefaultFontSize = 9 },
                new() { PlaceNumber = 13, FieldName = "ConsigneeDetails", Description = "13. Sender Instructions", DefaultX = 400f, DefaultY = 160f, DefaultFontSize = 8 },
                new() { PlaceNumber = 14, FieldName = "UnloadingInstructions", Description = "14. Instructions regarding Payment", DefaultX = 400f, DefaultY = 190f, DefaultFontSize = 8 },
                new() { PlaceNumber = 15, FieldName = "CarrierRemarks", Description = "15. Cash on Delivery", DefaultX = 40f, DefaultY = 430f, DefaultFontSize = 8 },
                new() { PlaceNumber = 16, FieldName = "ConsigneeRemarks", Description = "16. Carrier", DefaultX = 400f, DefaultY = 430f, DefaultFontSize = 8 },
                new() { PlaceNumber = 17, FieldName = "TransportAuthorizations", Description = "17. Successive Carriers", DefaultX = 40f, DefaultY = 500f, DefaultFontSize = 8 },
                new() { PlaceNumber = 18, FieldName = "RouteInfo", Description = "18. Carrier Observations", DefaultX = 200f, DefaultY = 500f, DefaultFontSize = 8 },
                new() { PlaceNumber = 19, FieldName = "InsuranceRemarks", Description = "19. Special Agreements", DefaultX = 40f, DefaultY = 540f, DefaultFontSize = 8 },
                new() { PlaceNumber = 20, FieldName = "CarrierSignature", Description = "20. To be Paid By", DefaultX = 40f, DefaultY = 580f, DefaultFontSize = 8 },
                new() { PlaceNumber = 21, FieldName = "ExportDate", Description = "21. Export/transport date", DefaultX = 400f, DefaultY = 540f, DefaultFontSize = 9 },
                new() { PlaceNumber = 22, FieldName = "SignaturePlace1", Description = "22. Signature Sender", DefaultX = 40f, DefaultY = 620f, DefaultFontSize = 8 },
                new() { PlaceNumber = 23, FieldName = "SignaturePlace2", Description = "23. Signature of the carrier", DefaultX = 270f, DefaultY = 620f, DefaultFontSize = 8 },
                new() { PlaceNumber = 24, FieldName = "SignaturePlace3", Description = "24. Signature Good received", DefaultX = 500f, DefaultY = 620f, DefaultFontSize = 8 },
            };
        }
    }
}
