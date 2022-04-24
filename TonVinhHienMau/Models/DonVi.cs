using System;
using System.Collections.Generic;

namespace TonVinhHienMau.Models
{
    public class DonVi
    {
        public Guid? Id { get; set; }
        public string Name { get; set; } 
        public string Code { get; set; }
        public string ParentId { get; set;}
        public bool IsDelete { get; set; }
        ICollection<NguoiHienMau> nguoiHienMaus { get; set; }
    }
}
