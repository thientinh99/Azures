//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WebApp.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Attendance
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Attendance()
        {
            this.AttendanceDetails = new HashSet<AttendanceDetail>();
        }
    
        public int AttID { get; set; }
        public Nullable<int> SubOfClaID { get; set; }
        public string StID { get; set; }
        public Nullable<int> Status { get; set; }
        public string Note { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AttendanceDetail> AttendanceDetails { get; set; }
        public virtual SubjectsOfClass SubjectsOfClass { get; set; }
        public virtual Student Student { get; set; }
    }
}
