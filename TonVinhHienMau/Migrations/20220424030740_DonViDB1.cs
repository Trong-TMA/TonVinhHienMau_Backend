using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace TonVinhHienMau.Migrations
{
    public partial class DonViDB1 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_NguoiHienMau_DonVis_DonViId",
                table: "NguoiHienMau");

            migrationBuilder.AlterColumn<Guid>(
                name: "DonViId",
                table: "NguoiHienMau",
                type: "uniqueidentifier",
                nullable: true,
                oldClrType: typeof(Guid),
                oldType: "uniqueidentifier");

            migrationBuilder.AddForeignKey(
                name: "FK_NguoiHienMau_DonVis_DonViId",
                table: "NguoiHienMau",
                column: "DonViId",
                principalTable: "DonVis",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_NguoiHienMau_DonVis_DonViId",
                table: "NguoiHienMau");

            migrationBuilder.AlterColumn<Guid>(
                name: "DonViId",
                table: "NguoiHienMau",
                type: "uniqueidentifier",
                nullable: false,
                defaultValue: new Guid("00000000-0000-0000-0000-000000000000"),
                oldClrType: typeof(Guid),
                oldType: "uniqueidentifier",
                oldNullable: true);

            migrationBuilder.AddForeignKey(
                name: "FK_NguoiHienMau_DonVis_DonViId",
                table: "NguoiHienMau",
                column: "DonViId",
                principalTable: "DonVis",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
