using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using TonVinhHienMau.Data;
using TonVinhHienMau.Models;
using TonVinhHienMau.Models.ViewModels;
using TonVinhHienMau.Services.Tools;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace TonVinhHienMau.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DonViController : ControllerBase
    {
        private readonly AppDbContext _context;
        private readonly ILogger<DonViController> _logger;

        public DonViController(AppDbContext context,
            ILogger<DonViController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet()]
        public IActionResult GetALL()
        {
            IEnumerable<DonVi> donvis = _context.DonVis.Where(u => u.IsDelete != true);
            return new JsonResult(donvis);
        }

        [HttpGet("GetById")]
        public IActionResult GetById(Guid DoViId)
        {
            DonVi donvis = _context.DonVis.FirstOrDefault(u => u.IsDelete != true && u.Id.Equals(DoViId));
            return new JsonResult(donvis);
        }

        [HttpPost("Create")]
        public IActionResult Create(DonViVm postData)
        {

            string DonViCode = RemoveUnicode.NonUnicode(postData.Name.ToUpper().Trim());
            DonViCode = RemoveUnicode.RemoveSpecialCrt(DonViCode);


            var donvi = _context.DonVis.FirstOrDefault(u => u.Code.Equals(DonViCode));

            if (donvi != null)
            {
                return new JsonResult(new { Message = "Đơn vị đã tồn tại" });
            }

            DonVi donVi = new DonVi()
            {
                Id = Guid.NewGuid(),
                Name = postData.Name,
                Code = DonViCode,
                ParentId = postData.ParentId,
                IsDelete = false,
            };
            _context.DonVis.Add(donVi);
            _context.SaveChanges();
            return new JsonResult("Tạo đơn vị thành công");
        }

        [HttpPost("Edit")]
        public IActionResult Edit(Guid id, DonViVm postData)
        {
            bool check = _context.DonVis.Any(u => u.Id.Equals(id));
            if (check)
            {
                var donvi = _context.DonVis.FirstOrDefault(u => u.Id.Equals(id));
                donvi.Name = postData.Name;
                donvi.ParentId = postData.ParentId;
                _context.DonVis.Update(donvi);
                _context.SaveChanges();
                return new JsonResult(new { Message = "Chỉnh sửa thành công" });
            }
            else
            {
                return new JsonResult(new { Message = "Đơn vị không tồn tại" });
            }
        }

        [HttpPut("Delete")]
        public IActionResult Delete(Guid id)
        {
            bool check = _context.DonVis.Any(u => u.Id.Equals(id));
            if (check)
            {
                var donvi = _context.DonVis.FirstOrDefault(u => u.Id.Equals(id));
                donvi.IsDelete = true;
                _context.DonVis.Update(donvi);
                _context.SaveChanges();
                return new JsonResult(new { Message = "Xóa đơn vị thành công" });
            }
            else
            {
                return new JsonResult(new { Message = "Đơn vị không tồn tại" });
            }
        }
    }
}
