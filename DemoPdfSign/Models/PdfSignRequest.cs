using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DemoPdfSign.Models
{
    public class PdfSignRequest
    {
        [Required(ErrorMessage = "Image data is not null")]
        public virtual string image { get; set; }
    }
}
