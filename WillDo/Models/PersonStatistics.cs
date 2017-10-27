using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WillDo.Models
{
    public class PersonSelections
    {
        [Required]
        public short? Question_1 { get; set; }
        [Required]
        public short? Question_2 { get; set; }
        [Required]
        public short? Question_3 { get; set; }
        [Required]
        public short? Question_4 { get; set; }
        [Required]
        public short? Question_5 { get; set; }
    }
}