using Microsoft.EntityFrameworkCore;
using DocStyleVerify.API.Models;

namespace DocStyleVerify.API.Data
{
    public class DocStyleVerifyDbContext : DbContext
    {
        public DocStyleVerifyDbContext(DbContextOptions<DocStyleVerifyDbContext> options) : base(options)
        {
        }

        public DbSet<Template> Templates { get; set; }
        public DbSet<TextStyle> TextStyles { get; set; }
        public DbSet<DirectFormatPattern> DirectFormatPatterns { get; set; }
        public DbSet<FormattingContext> FormattingContexts { get; set; }
        public DbSet<TabStop> TabStops { get; set; }
        public DbSet<VerificationResult> VerificationResults { get; set; }
        public DbSet<Mismatch> Mismatches { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Template configuration
            modelBuilder.Entity<Template>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Name).IsRequired().HasMaxLength(200);
                entity.Property(e => e.FileName).IsRequired().HasMaxLength(200);
                entity.Property(e => e.FilePath).IsRequired().HasMaxLength(500);
                entity.Property(e => e.FileHash).IsRequired().HasMaxLength(50);
                entity.Property(e => e.Status).IsRequired().HasMaxLength(20).HasDefaultValue("Active");
                entity.Property(e => e.CreatedBy).IsRequired().HasMaxLength(100);
                entity.Property(e => e.CreatedOn).IsRequired();
                entity.Property(e => e.Version).HasDefaultValue(1);
                
                entity.HasIndex(e => e.Name).IsUnique();
                entity.HasIndex(e => e.FileHash);
                entity.HasIndex(e => e.Status);
            });

            // TextStyle configuration
            modelBuilder.Entity<TextStyle>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Name).IsRequired().HasMaxLength(200);
                entity.Property(e => e.FontFamily).IsRequired().HasMaxLength(100);
                entity.Property(e => e.Color).IsRequired().HasMaxLength(10);
                entity.Property(e => e.Alignment).IsRequired().HasMaxLength(50);
                entity.Property(e => e.StyleType).IsRequired().HasMaxLength(50);
                entity.Property(e => e.StyleSignature).IsRequired().HasMaxLength(500);
                entity.Property(e => e.CreatedBy).IsRequired().HasMaxLength(100);
                entity.Property(e => e.CreatedOn).IsRequired();
                entity.Property(e => e.Version).HasDefaultValue(1);
                
                // Relationships
                entity.HasOne(e => e.Template)
                    .WithMany(t => t.TextStyles)
                    .HasForeignKey(e => e.TemplateId)
                    .OnDelete(DeleteBehavior.Cascade);

                entity.HasOne(e => e.FormattingContext)
                    .WithOne(f => f.TextStyle)
                    .HasForeignKey<TextStyle>(e => e.FormattingContextId)
                    .OnDelete(DeleteBehavior.Cascade);
                
                entity.HasIndex(e => e.TemplateId);
                entity.HasIndex(e => e.StyleSignature);
                entity.HasIndex(e => e.StyleType);
            });

            // DirectFormatPattern configuration
            modelBuilder.Entity<DirectFormatPattern>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.PatternName).IsRequired().HasMaxLength(200);
                entity.Property(e => e.Context).IsRequired().HasMaxLength(500);
                entity.Property(e => e.CreatedBy).IsRequired().HasMaxLength(100);
                entity.Property(e => e.CreatedOn).IsRequired();
                
                // Relationship
                entity.HasOne(e => e.TextStyle)
                    .WithMany(t => t.DirectFormatPatterns)
                    .HasForeignKey(e => e.TextStyleId)
                    .OnDelete(DeleteBehavior.Cascade);
                
                entity.HasIndex(e => e.TextStyleId);
                entity.HasIndex(e => e.Context);
            });

            // FormattingContext configuration
            modelBuilder.Entity<FormattingContext>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.StyleName).IsRequired().HasMaxLength(200);
                entity.Property(e => e.ElementType).IsRequired().HasMaxLength(100);
                entity.Property(e => e.ContextKey).HasMaxLength(500).HasColumnName("context_key");
                
                // JSON column for PostgreSQL
                entity.Property(e => e.ContentControlProperties)
                    .HasColumnType("jsonb")
                    .HasConversion(
                        v => System.Text.Json.JsonSerializer.Serialize(v, (System.Text.Json.JsonSerializerOptions?)null),
                        v => System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(v, (System.Text.Json.JsonSerializerOptions?)null) ?? new Dictionary<string, string>())
                    .Metadata.SetValueComparer(new Microsoft.EntityFrameworkCore.ChangeTracking.ValueComparer<Dictionary<string, string>>(
                        (c1, c2) => c1!.SequenceEqual(c2!),
                        c => c.Aggregate(0, (a, v) => HashCode.Combine(a, v.GetHashCode())),
                        c => c.ToDictionary(kvp => kvp.Key, kvp => kvp.Value)));
                
                entity.HasIndex(e => e.ContextKey);
                entity.HasIndex(e => e.ElementType);
                entity.HasIndex(e => e.StructuralRole);
            });

            // TabStop configuration
            modelBuilder.Entity<TabStop>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Alignment).IsRequired().HasMaxLength(20);
                entity.Property(e => e.Leader).IsRequired().HasMaxLength(20);
                
                // Relationship
                entity.HasOne(e => e.TextStyle)
                    .WithMany(t => t.TabStops)
                    .HasForeignKey(e => e.TextStyleId)
                    .OnDelete(DeleteBehavior.Cascade);
                
                entity.HasIndex(e => e.TextStyleId);
            });

            // VerificationResult configuration
            modelBuilder.Entity<VerificationResult>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.DocumentName).IsRequired().HasMaxLength(200);
                entity.Property(e => e.DocumentPath).IsRequired().HasMaxLength(500);
                entity.Property(e => e.Status).IsRequired().HasMaxLength(20).HasDefaultValue("Completed");
                entity.Property(e => e.CreatedBy).IsRequired().HasMaxLength(100);
                entity.Property(e => e.CreatedOn).IsRequired();
                
                // Relationship
                entity.HasOne(e => e.Template)
                    .WithMany()
                    .HasForeignKey(e => e.TemplateId)
                    .OnDelete(DeleteBehavior.Restrict);
                
                entity.HasIndex(e => e.TemplateId);
                entity.HasIndex(e => e.VerificationDate);
                entity.HasIndex(e => e.Status);
            });

            // Mismatch configuration
            modelBuilder.Entity<Mismatch>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.ContextKey).IsRequired().HasMaxLength(500);
                entity.Property(e => e.Location).IsRequired().HasMaxLength(1000);
                entity.Property(e => e.StructuralRole).IsRequired().HasMaxLength(100);
                entity.Property(e => e.MismatchFields).IsRequired().HasMaxLength(1000);
                entity.Property(e => e.Severity).IsRequired().HasMaxLength(50).HasDefaultValue("Medium");
                entity.Property(e => e.CreatedOn).IsRequired();
                
                // Relationship
                entity.HasOne(e => e.VerificationResult)
                    .WithMany(v => v.Mismatches)
                    .HasForeignKey(e => e.VerificationResultId)
                    .OnDelete(DeleteBehavior.Cascade);
                
                entity.HasIndex(e => e.VerificationResultId);
                entity.HasIndex(e => e.ContextKey);
                entity.HasIndex(e => e.Severity);
            });
        }
    }
} 