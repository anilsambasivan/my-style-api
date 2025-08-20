using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace DocStyleVerify.API.Migrations
{
    /// <inheritdoc />
    public partial class AlterColumnSizesTo5000 : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // Increase ErrorMessage field size in VerificationResults table
            migrationBuilder.AlterColumn<string>(
                name: "ErrorMessage",
                table: "VerificationResults",
                type: "character varying(5000)",
                maxLength: 5000,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(1000)",
                oldMaxLength: 1000);

            // Increase SampleText field size in FormattingContexts table
            migrationBuilder.AlterColumn<string>(
                name: "SampleText",
                table: "FormattingContexts",
                type: "character varying(5000)",
                maxLength: 5000,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(1000)",
                oldMaxLength: 1000);

            // Increase SampleText field size in DirectFormatPatterns table
            migrationBuilder.AlterColumn<string>(
                name: "SampleText",
                table: "DirectFormatPatterns",
                type: "character varying(5000)",
                maxLength: 5000,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(1000)",
                oldMaxLength: 1000);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            // Revert ErrorMessage field size in VerificationResults table
            migrationBuilder.AlterColumn<string>(
                name: "ErrorMessage",
                table: "VerificationResults",
                type: "character varying(1000)",
                maxLength: 1000,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(5000)",
                oldMaxLength: 5000);

            // Revert SampleText field size in FormattingContexts table
            migrationBuilder.AlterColumn<string>(
                name: "SampleText",
                table: "FormattingContexts",
                type: "character varying(1000)",
                maxLength: 1000,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(5000)",
                oldMaxLength: 5000);

            // Revert SampleText field size in DirectFormatPatterns table
            migrationBuilder.AlterColumn<string>(
                name: "SampleText",
                table: "DirectFormatPatterns",
                type: "character varying(1000)",
                maxLength: 1000,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "character varying(5000)",
                oldMaxLength: 5000);
        }
    }
}
