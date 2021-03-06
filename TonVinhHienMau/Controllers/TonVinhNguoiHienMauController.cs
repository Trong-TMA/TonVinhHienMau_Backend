using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using TonVinhHienMau.Data;
using TonVinhHienMau.Models;
using TonVinhHienMau.Models.ViewModels;

namespace TonVinhHienMau.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TonVinhNguoiHienMauController : ControllerBase
    {
        private readonly AppDbContext _context;
        private readonly ILogger<TonVinhNguoiHienMauController> _logger;

        public TonVinhNguoiHienMauController(AppDbContext context,
            ILogger<TonVinhNguoiHienMauController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet("GetALL")]
        public IActionResult getAll()
        {
            IEnumerable<NguoiHienMauVm> nguoihienmaus = _context.NguoiHienMau.Select(u => new NguoiHienMauVm()
            {
                HoTen = u.HoTen,
                DiaChi = u.DiaChi,
                GioiTinh = u.GioiTinh,
                NamSinh = u.NamSinh,
                NgheNghiep = u.NgheNghiep,
                TV_5 = u.TV_5,
                TV_10 = u.TV_10,
                TV_15 = u.TV_15,
                TV_20 = u.TV_20,
                TV_30 = u.TV_30,
                TV_40 = u.TV_40,
                TV_50 = u.TV_50,
                TV_60 = u.TV_60,
                TV_70 = u.TV_70,
                TV_80 = u.TV_80,
                TV_90 = u.TV_90,
                TV_100 = u.TV_100,
                NamTV_5 = u.NamTV_5,
                NamTV_10 = u.NamTV_10,
                NamTV_15 = u.NamTV_15,
                NamTV_20 = u.NamTV_20,
                NamTV_30 = u.NamTV_30,
                NamTV_40 = u.NamTV_40,
                NamTV_50 = u.NamTV_50,
                NamTV_60 = u.NamTV_60,
                NamTV_70 = u.NamTV_70,
                NamTV_80 = u.NamTV_80,
                NamTV_90 = u.NamTV_90,
                NamTV_100 = u.NamTV_100,
            });
            return new JsonResult(nguoihienmaus);
        }

        [HttpGet("GetByDonViId")]
        public IActionResult getByDonviId(Guid DonviId)
        {
            bool check = _context.DonVis.Any(u => u.Id.Equals(DonviId));
            if (check)
            {
                IEnumerable<NguoiHienMauVm> nguoihienmaus = _context.NguoiHienMau.Where(u => u.DonViId.Equals(DonviId)).Select(u => new NguoiHienMauVm()
                {
                    HoTen = u.HoTen,
                    DiaChi = u.DiaChi,
                    GioiTinh = u.GioiTinh,
                    NamSinh = u.NamSinh,
                    NgheNghiep = u.NgheNghiep,
                    TV_5 = u.TV_5,
                    TV_10 = u.TV_10,
                    TV_15 = u.TV_15,
                    TV_20 = u.TV_20,
                    TV_30 = u.TV_30,
                    TV_40 = u.TV_40,
                    TV_50 = u.TV_50,
                    TV_60 = u.TV_60,
                    TV_70 = u.TV_70,
                    TV_80 = u.TV_80,
                    TV_90 = u.TV_90,
                    TV_100 = u.TV_100,
                    NamTV_5 = u.NamTV_5,
                    NamTV_10 = u.NamTV_10,
                    NamTV_15 = u.NamTV_15,
                    NamTV_20 = u.NamTV_20,
                    NamTV_30 = u.NamTV_30,
                    NamTV_40 = u.NamTV_40,
                    NamTV_50 = u.NamTV_50,
                    NamTV_60 = u.NamTV_60,
                    NamTV_70 = u.NamTV_70,
                    NamTV_80 = u.NamTV_80,
                    NamTV_90 = u.NamTV_90,
                    NamTV_100 = u.NamTV_100,
                });
                return new JsonResult(nguoihienmaus);
            }
            else
            {
                return new JsonResult(new { Message = "Đơn vị không tồn tại" });
            }
        }

        [HttpPost("CheckHonor")]
        public IActionResult checkHonor(Guid? DonViId, IFormFile file)
        {
            List<NguoiHienMauVm> list_result = new List<NguoiHienMauVm>();
            if (file?.Length > 0)
            {
                // convert to a stream
                var stream = file.OpenReadStream();

                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.First();
                    var rowCount = worksheet.Dimension.Rows;
                    var DonVi = _context.DonVis.FirstOrDefault(u => u.Id.Equals(DonViId));
                    string NamTV = worksheet.Cells[2, 2].Value?.ToString().Trim();
                    for (var row = 3; row <= rowCount; row++)
                    {
                        string t5 = null, t10 = null, t15 = null, t20 = null, t30 = null, t40 = null, t50 = null, t60 = null, t70 = null, t80 = null, t90 = null, t100 = null;
                        var hoten = worksheet.Cells[row, 2].Value?.ToString().Trim();
                        var gioitinh = worksheet.Cells[row, 3].Value?.ToString().Trim();
                        var namsinh = worksheet.Cells[row, 4].Value?.ToString().Trim();
                        var nghenghiep = worksheet.Cells[row, 5].Value?.ToString().Trim();
                        var diachi = worksheet.Cells[row, 6].Value?.ToString().Trim();
                        var tv_5 = worksheet.Cells[row, 7].Value?.ToString().Trim();
                        var tv_10 = worksheet.Cells[row, 8].Value?.ToString().Trim();
                        var tv_15 = worksheet.Cells[row, 9].Value?.ToString().Trim();
                        var tv_20 = worksheet.Cells[row, 10].Value?.ToString().Trim();
                        var tv_30 = worksheet.Cells[row, 11].Value?.ToString().Trim();
                        var tv_40 = worksheet.Cells[row, 12].Value?.ToString().Trim();
                        var tv_50 = worksheet.Cells[row, 13].Value?.ToString().Trim();
                        var tv_60 = worksheet.Cells[row, 14].Value?.ToString().Trim();
                        var tv_70 = worksheet.Cells[row, 15].Value?.ToString().Trim();
                        var tv_80 = worksheet.Cells[row, 16].Value?.ToString().Trim();
                        var tv_90 = worksheet.Cells[row, 17].Value?.ToString().Trim();
                        var tv_100 = worksheet.Cells[row, 18].Value?.ToString().Trim();

                        if (tv_5 != null)
                        {
                            t5 = DonVi.Name;
                        }
                        else if (tv_10 != null)
                        {
                            t10 = DonVi.Name;
                        }
                        else if (tv_15 != null)
                        {
                            t15 = DonVi.Name;
                        }
                        else if (tv_20 != null)
                        {
                            t20 = DonVi.Name;
                        }
                        else if (tv_30 != null)
                        {
                            t30 = DonVi.Name;
                        }
                        else if (tv_40 != null)
                        {
                            t40 = DonVi.Name;
                        }
                        else if (tv_50 != null)
                        {
                            t50 = DonVi.Name;
                        }
                        else if (tv_60 != null)
                        {
                            t60 = DonVi.Name;
                        }
                        else if (tv_70 != null)
                        {
                            t70 = DonVi.Name;
                        }
                        else if (tv_80 != null)
                        {
                            t80 = DonVi.Name;
                        }
                        else if (tv_90 != null)
                        {
                            t90 = DonVi.Name;
                        }
                        else if (tv_100 != null)
                        {
                            t100 = DonVi.Name;
                        }

                        if (hoten != null && namsinh != null)
                        {
                            NguoiHienMauVm ng = new NguoiHienMauVm()
                            {

                                HoTen = hoten,
                                GioiTinh = Convert.ToBoolean(gioitinh),
                                NamSinh = Convert.ToInt32(namsinh),
                                NgheNghiep = nghenghiep,
                                DiaChi = diachi,
                                TV_5 = t5,
                                TV_10 = t10,
                                TV_15 = t15,
                                TV_20 = t20,
                                TV_30 = t30,
                                TV_40 = t40,
                                TV_50 = t50,
                                TV_60 = t60,
                                TV_70 = t70,
                                TV_80 = t80,
                                TV_90 = t90,
                                TV_100 = t100,
                                NamTV_5 = Convert.ToInt32(NamTV),
                                NamTV_10 = Convert.ToInt32(NamTV),
                                NamTV_15 = Convert.ToInt32(NamTV),
                                NamTV_20 = Convert.ToInt32(NamTV),
                                NamTV_30 = Convert.ToInt32(NamTV),
                                NamTV_40 = Convert.ToInt32(NamTV),
                                NamTV_50 = Convert.ToInt32(NamTV),
                                NamTV_60 = Convert.ToInt32(NamTV),
                                NamTV_70 = Convert.ToInt32(NamTV),
                                NamTV_80 = Convert.ToInt32(NamTV),
                                NamTV_90 = Convert.ToInt32(NamTV),
                                NamTV_100 = Convert.ToInt32(NamTV),
                                Note = null,

                            };
                            IEnumerable<NguoiHienMau> nguoiHienMau = _context.NguoiHienMau.Where(s => s.HoTen.Equals(ng.HoTen) && s.NamSinh == ng.NamSinh);

                            if (nguoiHienMau.Count() > 0)
                            {
                                foreach (var obj in nguoiHienMau)
                                {
                                    if (obj.TV_5 != null)
                                    {
                                        ng.Note = "Đã được tôn vinh mức 5" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_10 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 10" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_15 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 15" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_20 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 20" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_30 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 30" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_40 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 40" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_50 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 50" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_60 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 60" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_70 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 70" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_80 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 80" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_90 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 90" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                    else if (obj.TV_100 != null)
                                    {

                                        ng.Note = "Đã được tôn vinh mức 100" + " " + "Năm:" + NamTV;
                                        list_result.Add(ng);
                                    }
                                }
                            }

                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            return new JsonResult(list_result);
        }

        [HttpPost("ExportHonor")]
        public IActionResult ExportHonor(List<NguoiHienMauVm> list_result)
        {
            var stream = new MemoryStream();
            using (var xlPackage = new ExcelPackage(stream))
            {
                var worksheet = xlPackage.Workbook.Worksheets.Add("TonVinhHienMau");
                var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");
                namedStyle.Style.Font.UnderLine = true;
                namedStyle.Style.Font.Color.SetColor(Color.Blue);
                const int startRow = 5;
                var row = startRow;

                //Create Headers and format them
                worksheet.Cells["A1:D1"].Value = "UBND TỈNH BÌNH ĐỊNH";
                worksheet.Cells["A1:D1"].Merge = true;
                worksheet.Cells["A1:D1"].Style.Font.Bold = true;
                worksheet.Cells["A1:D1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;

                worksheet.Cells["A2:D2"].Value = "BCĐ VĐ HIẾN MÁU TÌNH NGUYỆN TỈNH";
                worksheet.Cells["A2:D2"].Merge = true;
                worksheet.Cells["A2:D2"].Style.Font.Bold = true;
                worksheet.Cells["A2:D2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;

                worksheet.Cells["A3:S3"].Value = "DANH SÁCH CÁ NHÂN HIẾN MÁU TÌNH NGUYỆN 5 LẦN TRỞ LÊN CHƯA ĐƯỢC TÔN VINH TỈNH BÌNH ĐỊNH NĂM ..." +"\n"+
                    "(Căn cứ theo Quy chế Tôn vinh, khen thưởng cá nhân, tập thể có thành tích Hiến máu tình nguyện và vận động hiến máu tình nguyện tại Quyết định số 122/QĐ-BCĐQG ngày ... tháng ... năm ... của Ban Chỉ đạo quốc gia vận động hiến máu tình nguyện)";
                worksheet.Cells["A3:S3"].Merge = true;
                worksheet.Cells["A3:S3"].Style.Font.Bold = true;
                worksheet.Cells["A3:S3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;

                worksheet.Cells["E1:G1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                worksheet.Cells["E1:G1"].Merge = true;
                worksheet.Cells["E1:G1"].Style.Font.Bold = true;
                worksheet.Cells["E1:G1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;

                worksheet.Cells["E2:G2"].Value = "Độc lập - Tự do - Hạnh phúc";
                worksheet.Cells["E2:G2"].Merge = true;
                worksheet.Cells["E2:G2"].Style.Font.Bold = true;
                worksheet.Cells["E2:G2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;

                worksheet.Cells["A4"].Value = "STT";
                worksheet.Cells["B4"].Value = "Họ tên";
                worksheet.Cells["C4"].Value = "Giới tính";
                worksheet.Cells["D4"].Value = "Năm Sinh";
                worksheet.Cells["E4"].Value = "Nghề Nghiệp";
                worksheet.Cells["F4"].Value = "Địa chỉ";
                worksheet.Cells["G4"].Value = "5 Lần";
                worksheet.Cells["H4"].Value = "10 Lần";
                worksheet.Cells["I4"].Value = "15 Lần";
                worksheet.Cells["J4"].Value = "20 Lần";
                worksheet.Cells["K4"].Value = "30 Lần";
                worksheet.Cells["L4"].Value = "40 Lần";
                worksheet.Cells["M4"].Value = "50 Lần";
                worksheet.Cells["N4"].Value = "60 Lần";
                worksheet.Cells["O4"].Value = "70 Lần";
                worksheet.Cells["P4"].Value = "80 Lần";
                worksheet.Cells["Q4"].Value = "90 Lần";
                worksheet.Cells["R4"].Value = "100 Lần";
                worksheet.Cells["S4"].Value = "Lưu ý";
                

                worksheet.Cells["A4:F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A4:F4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                worksheet.Cells["A4:F4"].Style.Font.Bold = true;

                row = 5;
                int count =  0;
                foreach (var item in list_result)
                {
                    worksheet.Cells[row, 1].Value = count + 1;
                    worksheet.Cells[row, 2].Value = item.HoTen;
                    worksheet.Cells[row, 3].Value = item.GioiTinh;
                    worksheet.Cells[row, 4].Value = item.NamSinh;
                    worksheet.Cells[row, 5].Value = item.NgheNghiep;
                    worksheet.Cells[row, 6].Value = item.DiaChi;
                    worksheet.Cells[row, 7].Value = item.TV_5;
                    worksheet.Cells[row, 8].Value = item.TV_10;
                    worksheet.Cells[row, 9].Value = item.TV_15;
                    worksheet.Cells[row, 10].Value = item.TV_20;
                    worksheet.Cells[row, 11].Value = item.TV_30;
                    worksheet.Cells[row, 12].Value = item.TV_40;
                    worksheet.Cells[row, 13].Value = item.TV_50;
                    worksheet.Cells[row, 14].Value = item.TV_60;
                    worksheet.Cells[row, 15].Value = item.TV_70;
                    worksheet.Cells[row, 16].Value = item.TV_80;
                    worksheet.Cells[row, 17].Value = item.TV_90;
                    worksheet.Cells[row, 18].Value = item.TV_100;
                    worksheet.Cells[row, 19].Value = item.Note;

                    row++;
                }

                // set some core property values
                xlPackage.Workbook.Properties.Title = "Danh sách người hiến máu";
               /* xlPackage.Workbook.Properties.Author = "Mohamad Lawand";
                xlPackage.Workbook.Properties.Subject = "User List";*/
                // save the new spreadsheet
                xlPackage.Save();
                // Response.Clear();
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TonVinhHienMau.xlsx");
        }
    }
}
