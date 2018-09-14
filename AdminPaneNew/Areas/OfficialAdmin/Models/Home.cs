using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace AdminPaneNew.Areas.OfficialAdmin.Models
{
    public class Home
    {

    }
    public class slider
    {
        [Key]
        public int Sliderid { get; set; }
        [Required]
        public string Name { get; set; }
        [AllowHtml]
        public string Description { get; set; }
        public string Image { get; set; }
        public string url { get; set; }

        public DateTime date { get; set; }

    }
    public class news
    {
        [Key]
        public int Newsid { get; set; }
        [Required]
        public string Name { get; set; }
        [AllowHtml]
        public string Description { get; set; }
        [AllowHtml]
        public string ShortDescription { get; set; }
        public string Image { get; set; }
        public string Thumbnail { get; set; }
        public DateTime date { get; set; }
    }
    public class Service
    {
        [Key]
        public int Serviceid { get; set; }
        [Required]
        public string Name { get; set; }
        [AllowHtml]
        public string Description { get; set; }
        [AllowHtml]
        public string ShortDescription { get; set; }
        public string Image { get; set; }

        public DateTime date { get; set; }
        public string Thumbnail { get; set; }
        public string Keyword { get; set; }
        public string MetaDescription { get; set; }
    }
    public class Testimonial
    {
        [Key]
        public int Testid { get; set; }
        [Required]
        public string Name { get; set; }
        [AllowHtml]
        public string Description { get; set; }
        public string Image { get; set; }
        public DateTime date { get; set; }
    }
    public class Album
    {
        [Key]
        public int Albumid { get; set; }
        public string AlbumName { get; set; }
        public string Image { get; set; }
        public DateTime date { get; set; }
        public virtual ICollection<Gallery> Galleries { get; set; }

    }
    public class Gallery
    {
        [Key]
        public int GalleryId { get; set; }
        [DisplayName("Album")]
        public int Albumid { get; set; }

        public virtual Album Albums { get; set; }
        public string Thumbnail { get; set; }
        public string Images { get; set; }
        public DateTime date { get; set; }
    }
    public class Clientlogo
    {
        [Key]
        public int Clientid { get; set; }
        [Required]
        public string Name { get; set; }
        public string Image { get; set; }
        public DateTime date { get; set; }

    }
    public class Logo
    {
        [Key]
        public int Logoid { get; set; }
        [Required]
        public string Name { get; set; }
        public string Image { get; set; }
        public DateTime date { get; set; }

    }
    public class Contact
    {
        [Key]
        public int Contactid { get; set; }
        [Required]
        public string Name { get; set; }
        public string Phone { get; set; }
        public string Mobile { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public DateTime date { get; set; }
    }
    public class Features
    {
        [Key]
        public int Featureid { get; set; }
        [Required]
        public string Name { get; set; }
        [AllowHtml]
        public string Description { get; set; }
        [AllowHtml]
        public string ShortDescription { get; set; }
        public string Image { get; set; }

        public DateTime date { get; set; }
        public string Thumbnail { get; set; }
        public string Keyword { get; set; }
        public string MetaDescription { get; set; }
    }
    public class Videos
    {
        [Key]
        public int Videoid { get; set; }
        [Required]
        public string Name { get; set; }
        public string Image { get; set; }
        public string Url { get; set; }
        public DateTime date { get; set; }

    }
    public class Account
    {
        [Key]
        public int Accountid { get; set; }
        [Required]
        public string Name { get; set; }
        public string Mobile { get; set; }
        public string Usename { get; set; }
        public string Password { get; set; }
        public string Role { get; set; }
        public DateTime date { get; set; }
    }
    public class StudentDetail
    {
        [Key]
        public int Studentid { get; set; }
        [Required]
        public string Name { get; set; }
        public string FatherName { get; set; }
        [RegularExpression(@"^([0-9]{10})$", ErrorMessage = "Invalid Mobile Number.")]
        public string Mobile { get; set; }
        [DataType(DataType.EmailAddress, ErrorMessage = "Please enter a valid email address.")]
        [RegularExpression("^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$", ErrorMessage = "You must provide a valid email address.")]
        public string Email { get; set; }
        public string Gender { get; set; }
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public Nullable<System.DateTime> DOB { get; set; }
        public string Image { get; set; }
        public string Address { get; set; }
        public string Qualification { get; set; }
        public string CourseType { get; set; }
        public string RollNo { get; set; }
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy}", ApplyFormatInEditMode = true)]
        public Nullable<System.DateTime> JoiningDate { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public DateTime date { get; set; }
        public Byte Status { get; set; }
        public virtual ICollection<AssignTest> AssignTest { get; set; }
    }
    public class Category
    {
        [Key]
        public int Categoryid { get; set; }
        public string Name { get; set; }
        public DateTime date { get; set; }
        public virtual ICollection<IeltsTest> IeltsTest { get; set; }
    }
    public enum Gender
    {
        Male,
        Female
    }
    public class IeltsTest
    {
        [Key]
        public int Ieltsid { get; set; }
        [Required]
        public string Name { get; set; }
        public int Categoryid { get; set; }
        public virtual Category Category { get; set; }
        public string TestType { get; set; }
        public string Image { get; set; }
        public string Url { get; set; }
        public string Audio { get; set; }
        public DateTime date { get; set; }
        public virtual ICollection<AssignTest> AssignTest { get; set; }
    }
    public enum TestType
    {
        General,
        Academic

    }
    public class AssignTest
    {
        [Key]
        public int Assignid { get; set; }
        public int Studentid { get; set; }
        public virtual StudentDetail StudentDetail { get; set; }
        public int Ieltsid { get; set; }
        public virtual IeltsTest IeltsTest { get; set; }
        public string Image { get; set; }
        public DateTime date { get; set; }
        public string Status { get; set; }
    }

    public class StudentReg
    {
        [Key]
        public int Studentid { get; set; }
        [Required]
        public string StudentName { get; set; }
        public string FatherName { get; set; }
        public string Address { get; set; }
        [RegularExpression(@"^([0-9]{10})$", ErrorMessage = "Invalid Mobile Number.")]
        public string Contact { get; set; }
        public string Laststudy { get; set; }
        public string Medical { get; set; }
        public string Refusal { get; set; }
        [DataType(DataType.EmailAddress, ErrorMessage = "Please enter a valid email address.")]
        [RegularExpression("^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$", ErrorMessage = "You must provide a valid email address.")]
        public string Email { get; set; }
        public string Password { get; set; }
        public int RollNo { get; set; }
        public string Fileno { get; set; }

    }
    public class fees
    {
        [Key]
        public int feeid { get; set; }
        public string studentid { get; set; }
        public int Package { get; set; }
        public int Advance { get; set; }
        public int pay { get; set; }
        public int balance { get; set; }
    }
    public class SingleFee
    {
        [Key]
        public int sfid { get; set; }
        public string studentid { get; set; }
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy}", ApplyFormatInEditMode = true)]
        public Nullable<System.DateTime> Date { get; set; }
        public int Paid { get; set; }
        public int Billno { get; set; }
        public string Receivedby { get; set; }

    }
}