using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using TonVinhHienMau.Data;
using TonVinhHienMau.Models;
using TonVinhHienMau.Models.ViewModels;

namespace TonVinhHienMau.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ThemNguoiHienMauController : ControllerBase
    {
        private readonly AppDbContext _context;
        private readonly ILogger<ThemNguoiHienMauController> _logger;
        public ThemNguoiHienMauController(AppDbContext context,
            ILogger<ThemNguoiHienMauController> logger)
        {
            _context = context;
            _logger = logger;
        }


        [HttpPost("AddFromExceltoData")]
        public IActionResult importPeople(Guid DonViId,IFormFile file)
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


                    var DonVi = _context.DonVis.FirstOrDefault(u=>u.Id.Equals(DonViId));
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
                            NguoiHienMau nguoiHienMau = new NguoiHienMau()
                            {
                                DonViId = DonVi.Id,
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
                            };
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
                            var nguoiHM = _context.NguoiHienMau.FirstOrDefault(s => s.HoTen.Equals(nguoiHienMau.HoTen) && s.NamSinh == nguoiHienMau.NamSinh);

                            if (nguoiHM == null)
                            {
                                _context.NguoiHienMau.Add(nguoiHienMau);
                                ng.Note = "Đã thêm mới";
                                list_result.Add(ng);

                            }
                            else
                            {
                                nguoiHienMau.NgheNghiep = ng.NgheNghiep;
                                nguoiHienMau.DiaChi = ng.DiaChi;
                                nguoiHienMau.TV_5 = ng.TV_5;
                                nguoiHienMau.TV_10 = ng.TV_10;
                                nguoiHienMau.TV_15 = ng.TV_15;
                                nguoiHienMau.TV_20 = ng.TV_20;
                                nguoiHienMau.TV_30 = ng.TV_30;
                                nguoiHienMau.TV_40 = ng.TV_40;
                                nguoiHienMau.TV_50 = ng.TV_50;
                                nguoiHienMau.TV_60 = ng.TV_60;
                                nguoiHienMau.TV_70 = ng.TV_70;
                                nguoiHienMau.TV_80 = ng.TV_80;
                                nguoiHienMau.TV_90 = ng.TV_90;
                                nguoiHienMau.TV_100 = ng.TV_100;

                                nguoiHienMau.NamTV_5 = ng.NamTV_5;
                                nguoiHienMau.NamTV_10 = ng.NamTV_10;
                                nguoiHienMau.NamTV_15 = ng.NamTV_15;
                                nguoiHienMau.NamTV_20 = ng.NamTV_20;
                                nguoiHienMau.NamTV_30 = ng.NamTV_30;
                                nguoiHienMau.NamTV_40 = ng.NamTV_40;
                                nguoiHienMau.NamTV_50 = ng.NamTV_50;
                                nguoiHienMau.NamTV_60 = ng.NamTV_60;
                                nguoiHienMau.NamTV_70 = ng.NamTV_70;
                                nguoiHienMau.NamTV_80 = ng.NamTV_80;
                                nguoiHienMau.NamTV_90 = ng.NamTV_90;
                                nguoiHienMau.NamTV_100 = ng.NamTV_100;

                                _context.NguoiHienMau.Update(nguoiHienMau);
                                ng.Note = "Đã tồn tại và được cập nhật";
                                list_result.Add(ng);
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                    _context.SaveChanges();
                }

           }
           return new JsonResult(list_result);
        }


    }
}
