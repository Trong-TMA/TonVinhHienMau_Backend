using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace TonVinhHienMau.Migrations
{
    public partial class DonViDB : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "DonVis",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    Name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Code = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ParentId = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    IsDelete = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DonVis", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "NguoiHienMau",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DonViId = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    HoTen = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    GioiTinh = table.Column<bool>(type: "bit", nullable: false),
                    NamSinh = table.Column<int>(type: "int", nullable: false),
                    NgheNghiep = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DiaChi = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    TV_5 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_5 = table.Column<int>(type: "int", nullable: false),
                    TV_10 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_10 = table.Column<int>(type: "int", nullable: false),
                    TV_15 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_15 = table.Column<int>(type: "int", nullable: false),
                    TV_20 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_20 = table.Column<int>(type: "int", nullable: false),
                    TV_30 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_30 = table.Column<int>(type: "int", nullable: false),
                    TV_40 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_40 = table.Column<int>(type: "int", nullable: false),
                    TV_50 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_50 = table.Column<int>(type: "int", nullable: false),
                    TV_60 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_60 = table.Column<int>(type: "int", nullable: false),
                    TV_70 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_70 = table.Column<int>(type: "int", nullable: false),
                    TV_80 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_80 = table.Column<int>(type: "int", nullable: false),
                    TV_90 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_90 = table.Column<int>(type: "int", nullable: false),
                    TV_100 = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NamTV_100 = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NguoiHienMau", x => x.Id);
                    table.ForeignKey(
                        name: "FK_NguoiHienMau_DonVis_DonViId",
                        column: x => x.DonViId,
                        principalTable: "DonVis",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_NguoiHienMau_DonViId",
                table: "NguoiHienMau",
                column: "DonViId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "NguoiHienMau");

            migrationBuilder.DropTable(
                name: "DonVis");
        }
    }
}
