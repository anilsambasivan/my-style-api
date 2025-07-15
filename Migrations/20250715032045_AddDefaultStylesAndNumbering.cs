using System;
using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace DocStyleVerify.API.Migrations
{
    /// <inheritdoc />
    public partial class AddDefaultStylesAndNumbering : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "DefaultStyles",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    TemplateId = table.Column<int>(type: "integer", nullable: false),
                    StyleId = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false),
                    Name = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false),
                    Type = table.Column<string>(type: "character varying(50)", maxLength: 50, nullable: false),
                    BasedOn = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false),
                    NextStyle = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false),
                    IsDefault = table.Column<bool>(type: "boolean", nullable: false),
                    IsCustom = table.Column<bool>(type: "boolean", nullable: false),
                    Priority = table.Column<int>(type: "integer", nullable: false),
                    IsHidden = table.Column<bool>(type: "boolean", nullable: false),
                    IsQuickStyle = table.Column<bool>(type: "boolean", nullable: false),
                    FontFamily = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    FontSize = table.Column<float>(type: "real", nullable: false),
                    IsBold = table.Column<bool>(type: "boolean", nullable: false),
                    IsItalic = table.Column<bool>(type: "boolean", nullable: false),
                    IsUnderline = table.Column<bool>(type: "boolean", nullable: false),
                    Color = table.Column<string>(type: "character varying(10)", maxLength: 10, nullable: false),
                    Alignment = table.Column<string>(type: "character varying(50)", maxLength: 50, nullable: false),
                    SpacingBefore = table.Column<float>(type: "real", nullable: false),
                    SpacingAfter = table.Column<float>(type: "real", nullable: false),
                    IndentationLeft = table.Column<float>(type: "real", nullable: false),
                    IndentationRight = table.Column<float>(type: "real", nullable: false),
                    FirstLineIndent = table.Column<float>(type: "real", nullable: false),
                    LineSpacing = table.Column<float>(type: "real", nullable: false),
                    RawXml = table.Column<string>(type: "text", nullable: false),
                    CreatedBy = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    CreatedOn = table.Column<DateTime>(type: "timestamp with time zone", nullable: false),
                    ModifiedBy = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    ModifiedOn = table.Column<DateTime>(type: "timestamp with time zone", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DefaultStyles", x => x.Id);
                    table.ForeignKey(
                        name: "FK_DefaultStyles_Templates_TemplateId",
                        column: x => x.TemplateId,
                        principalTable: "Templates",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "NumberingDefinitions",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    TemplateId = table.Column<int>(type: "integer", nullable: false),
                    AbstractNumId = table.Column<int>(type: "integer", nullable: false),
                    NumberingId = table.Column<int>(type: "integer", nullable: false),
                    Name = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false),
                    Type = table.Column<string>(type: "character varying(50)", maxLength: 50, nullable: false),
                    RawXml = table.Column<string>(type: "text", nullable: false),
                    CreatedBy = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    CreatedOn = table.Column<DateTime>(type: "timestamp with time zone", nullable: false),
                    ModifiedBy = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    ModifiedOn = table.Column<DateTime>(type: "timestamp with time zone", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NumberingDefinitions", x => x.Id);
                    table.ForeignKey(
                        name: "FK_NumberingDefinitions_Templates_TemplateId",
                        column: x => x.TemplateId,
                        principalTable: "Templates",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "NumberingLevels",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    NumberingDefinitionId = table.Column<int>(type: "integer", nullable: false),
                    Level = table.Column<int>(type: "integer", nullable: false),
                    NumberFormat = table.Column<string>(type: "character varying(50)", maxLength: 50, nullable: false),
                    LevelText = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false),
                    LevelJustification = table.Column<string>(type: "character varying(10)", maxLength: 10, nullable: false),
                    StartValue = table.Column<int>(type: "integer", nullable: false),
                    IsLegal = table.Column<bool>(type: "boolean", nullable: false),
                    FontFamily = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    FontSize = table.Column<float>(type: "real", nullable: false),
                    IsBold = table.Column<bool>(type: "boolean", nullable: false),
                    IsItalic = table.Column<bool>(type: "boolean", nullable: false),
                    Color = table.Column<string>(type: "character varying(10)", maxLength: 10, nullable: false),
                    IndentationLeft = table.Column<float>(type: "real", nullable: false),
                    IndentationHanging = table.Column<float>(type: "real", nullable: false),
                    TabStopPosition = table.Column<float>(type: "real", nullable: false),
                    RawXml = table.Column<string>(type: "text", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NumberingLevels", x => x.Id);
                    table.ForeignKey(
                        name: "FK_NumberingLevels_NumberingDefinitions_NumberingDefinitionId",
                        column: x => x.NumberingDefinitionId,
                        principalTable: "NumberingDefinitions",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_DefaultStyles_StyleId",
                table: "DefaultStyles",
                column: "StyleId");

            migrationBuilder.CreateIndex(
                name: "IX_DefaultStyles_TemplateId",
                table: "DefaultStyles",
                column: "TemplateId");

            migrationBuilder.CreateIndex(
                name: "IX_DefaultStyles_Type",
                table: "DefaultStyles",
                column: "Type");

            migrationBuilder.CreateIndex(
                name: "IX_NumberingDefinitions_AbstractNumId",
                table: "NumberingDefinitions",
                column: "AbstractNumId");

            migrationBuilder.CreateIndex(
                name: "IX_NumberingDefinitions_NumberingId",
                table: "NumberingDefinitions",
                column: "NumberingId");

            migrationBuilder.CreateIndex(
                name: "IX_NumberingDefinitions_TemplateId",
                table: "NumberingDefinitions",
                column: "TemplateId");

            migrationBuilder.CreateIndex(
                name: "IX_NumberingLevels_Level",
                table: "NumberingLevels",
                column: "Level");

            migrationBuilder.CreateIndex(
                name: "IX_NumberingLevels_NumberingDefinitionId",
                table: "NumberingLevels",
                column: "NumberingDefinitionId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "DefaultStyles");

            migrationBuilder.DropTable(
                name: "NumberingLevels");

            migrationBuilder.DropTable(
                name: "NumberingDefinitions");
        }
    }
}
