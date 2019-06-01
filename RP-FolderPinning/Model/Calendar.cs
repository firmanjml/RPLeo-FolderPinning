using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RP_FolderPinning.Model
{
    class Calendar
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public double StartTime { get; set; }
        public string StartTimeStr { get; set; }
        public double EndTime { get; set; }
        public string EndTimeStr { get; set; }
        public int Duration { get; set; }
        public bool IsAllDayEvent { get; set; }
        public string Organizer { get; set; }
        public string OrganizerId { get; set; }
        public string Organizers { get; set; }
        public string Venue { get; set; }
        public bool IsRequiredAttendance { get; set; }
        public string Description { get; set; }
        public int Type { get; set; }
        public string EventTypeStr { get; set; }
        public bool IsStudent { get; set; }
        public List<string> Participants { get; set; }
        public bool IsEditEnable { get; set; }
        public bool IgnoreAttendanceChanges { get; set; }
        public bool IsTakePerAttendance { get; set; }
        public string LessonNo { get; set; }
        public string LessonName { get; set; }
        public bool EnableAttendanceTaking { get; set; }
        public string LessonURL { get; set; }
        public string SurveyURL { get; set; }
        public string SurveyStatus { get; set; }
        public string MUEModuleId { get; set; }
        public string MUESchoolName { get; set; }
        public string LearningPathId { get; set; }
        public string LearningObjectId { get; set; }
        public int LOType { get; set; }
        public string MClassId { get; set; }

    }
}
