namespace STL.Models {
    public class RentalRate {
        public string MatNo { get; set; }
        public double Rate { get; set; }
    }

    public class PlanningMaterial : RentalRate {
        public int Qty { get; set; }
        public string Description { get; set; }
        public string Package { get; set; }
        public string DeliveryDestination { get; set; }

        public string EqmOwnership { get; set; }

        public string Region { get; set; }
        public string PackageId { get; set; }
    }

    public class DeliveryDestination {
        public string Name { get; set; }
        public string PersonName { get; set; }
        public string Address { get; set; }
        public string WBS { get; set; }
        public string FOPO { get; set; }
        public string ExecutionPlant { get; set; }
        public string Incoterm { get; set; }
        public string IncotermLocation { get; set; }
    }

    public class Source {

        public string Name { get; set; }
        public string EQMProvider { get; set; }
        public string Region { get; set; }
    }
}
