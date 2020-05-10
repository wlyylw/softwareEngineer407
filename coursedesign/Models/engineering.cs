using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace coursedesign
{
    [MetadataType(typeof(EnMetaData))]
    public partial class engineering
    {
        class EnMetaData
        {

            public Nullable<System.DateTime> birth { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "*请填写性别")]
            [RegularExpression(@"^[0-1]$", ErrorMessage = "*请输入正确的性别编号")]
            public int sex { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "请填写姓名")]
            [StringLength(20, ErrorMessage = "*姓名长度应小于20个字符")]
            public string name { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "*请填写员工编号")]
            [Display(Name = "员工编号")]
            [RegularExpression(@"^[0-9]{1,4}$", ErrorMessage = "*请输入正确的员工编号，例：0001")]
            public int id { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "*请填写籍贯")]
            [StringLength(10, ErrorMessage = "*籍贯格式填写错误")]
            public string place { get; set; }

            [Required(AllowEmptyStrings = false,ErrorMessage = "*请填写学历")]
            [Display(Name = "学历")]
            [RegularExpression(@"^[0-4]$", ErrorMessage = "*请输入正确的学历编号")]
            public Nullable<int> education { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "*请填写地址")]
            [StringLength(10, ErrorMessage = "*地址格式填写错误")]
            public string address { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "*请填写电话")]
            [RegularExpression(@"^1[0-9]{9}$", ErrorMessage = "*电话格式错误")]
            public string telephone { get; set; }

            [Required(AllowEmptyStrings = false, ErrorMessage = "*请填写工龄")]
            [RegularExpression(@"^(([1-4][0-9]?)|50)$", ErrorMessage = "*工龄超限")]
            public string workage { get; set; }

            [RegularExpression(@"^[^0]$", ErrorMessage = "*工资错误")]
            public int salary { get; set; }
        }

    }
}