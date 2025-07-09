using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.Models
{
    public class TabStop
    {
        [Key]
        public int Id { get; set; }
        
        public int TextStyleId { get; set; }
        
        [Required]
        public float Position { get; set; }
        
        [Required]
        [StringLength(20)]
        public string Alignment { get; set; } = string.Empty; // Left, Right, Center, Decimal, Bar
        
        [Required]
        [StringLength(20)]
        public string Leader { get; set; } = string.Empty; // None, Dots, Dashes, Lines, Heavy, MiddleDot
        
        // Navigation property
        public virtual TextStyle TextStyle { get; set; } = null!;
    }
} 