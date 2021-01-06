using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Extension.Excel.Impl.Import.Tests
{
    internal class SalaryRecordImportModel
    {
        /// <summary>
        /// 姓名
        /// </summary>
        [Description("姓名")]
        [Required(ErrorMessage = "{0}不能为空")]
        public string UserName { get; set; }
        /// <summary>
        /// 手机号
        /// </summary>
        [Description("手机号")]
        [Required]
        public string Phone { get; set; }
        /// <summary>
        /// 身份证
        /// </summary>
        [Description("身份证号")]
        public string IDCard { get; set; }

        /// <summary>
        /// 企业
        /// </summary>
        [Description("企业名称")]
        [Required(ErrorMessage = "企业名称不能为空")]
        public string EntName { get; set; }
        /// <summary>
        /// 部门
        /// </summary>
        [Description("部门")]
        public string DeptName { get; set; }
        /// <summary>
        /// 职位
        /// </summary>
        [Description("职位")]
        public string PositionName { get; set; }
        /// <summary>
		/// 用人单位
		/// </summary>
        [Description("用人单位")]
        public string EmployerName { get; set; }
        /// <summary>
        /// 银行卡名称
        /// </summary>
        [Description("开户银行")]
        public string BankName { get; set; }

        /// <summary>
        /// 开户地
        /// </summary>
        [Description("开户地")]
        public string BankAddress { get; set; }

        /// <summary>
        /// 开户支行
        /// </summary>
        [Description("开户支行")]
        public string BankBranch { get; set; }
        /// <summary>
        /// 银行卡号
        /// </summary>
        [Description("银行账号")]
        public string BankCard { get; set; }

        /// <summary>
        /// 应发工资
        /// </summary>
        [Description("应发工资")]
        public decimal WagesPayable { get; set; }

        /// <summary>
        /// 扣款
        /// </summary>
        [Description("扣款")]
        public decimal Deduction { get; set; }

        /// <summary>
        /// 实发工资
        /// </summary>
        [Description("实发工资")]
        public decimal RealPay { get; set; }

        [Description("百分比")]
        public decimal Percent { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        [Description("备注")]
        public string Remark { get; set; }
    }
}