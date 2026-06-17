using System.Collections.Generic;

namespace CMRPrint
{
    public class FieldAssignment
    {
        public string FieldName { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    public class ProfileRecord
    {
        public string Name { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Place { get; set; } = string.Empty;
        public List<FieldAssignment> FieldAssignments { get; set; } = new();

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Name) ? "(new profile)" : Name;
        }
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string VatNumber { get; set; } = string.Empty;
        public string ExporterProfileName { get; set; } = string.Empty;
        public string TransportProfileName { get; set; } = string.Empty;
        public string LoadingPlaceProfileName { get; set; } = string.Empty;
        public string PlaceOfIssue { get; set; } = string.Empty;
        public List<FieldAssignment> FieldAssignments { get; set; } = new();

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Name) ? "(new customer)" : Name;
        }
    }
}
