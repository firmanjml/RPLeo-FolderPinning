# RPLeo-FolderPinning

## What is this for?
This Program allows you to pin your class folder automatically base on the calendar data that was published in LEO to your taskbar everytime you start up your pc.

The calendar data must be retrieved manually from L.E.O via RetrieveEvents post request in Calendar.

## Installation & Program configuration

RPLeo-FolderPinning requires [Syspin.exe](http://www.technosys.net/products/utils/pintotaskbar) to run.

Download the program and place it in the same directory where the program is.

Make config.json and put it in the same directory.

Example of the config.json can be found below.
```sh
{
  "calendar": [
    {
      "Id": "xxxxxxx-xxxxx-xxxx-xxx-xxxxxx",
      "Title": "C322-4F-E61H-A",
      "StartTime": 1556761500000,
      "StartTimeStr": null,
      "EndTime": 1556785800000,
      "EndTimeStr": null,
      "Duration": 0,
      "IsAllDayEvent": false,
      "Organizer": null,
      "OrganizerId": null,
      "Organizers": null,
      "Venue": "E61H",
      "IsRequiredAttendance": false,
      "Description": null,
      "Type": 1,
      "EventTypeStr": "Lesson",
      "IsStudent": false,
      "Participants": [],
      "IsEditEnable": false,
      "IgnoreAttendanceChanges": false,
      "IsTakePerAttendance": false,
      "LessonNo": null,
      "LessonName": null,
      "EnableAttendanceTaking": false,
      "LessonURL": "/ClassManagement/Lesson/LessonDetails?classId=xxxxxxx-xxxxx-xxxx-xxx-xxxxxx&lessonId=",
      "SurveyURL": null,
      "SurveyStatus": null,
      "MUEModuleId": "xxxxxxx-xxxxx-xxxx-xxx-xxxxxx",
      "MUESchoolName": null,
      "LearningPathId": null,
      "LearningObjectId": null,
      "LOType": 0,
      "MClassId": "xxxxxxx-xxxxx-xxxx-xxx-xxxxxx"
    }
  ],
  "prev_dir": "",
  "folder_dir": "D:\\Schools\\Semester 1",
  "verbose": true,
  "verbose_sleep_time": 3
}
````

### Attributes of the configuration
| Key | Type | Info | Mandatory
| ------ | ------ | ------ | ------ |
| calendar | Array of Objects | Information of your calendar | tue
| prev_dir | String | Cache directory of current lesson | false
| folder_dir | String | The directory of your lessons folder | true
| verbose | Boolean | Show more more information | true
| verbose_sleep_time | Integer | Delay of verbose mode | true


### Folder structuring
Example of how your directory of your module should look like
![alt text](https://i.imgur.com/TQOMwQc.jpg)
Alternatively, you can just name the folders as the Module Code.
The program will automatically create a short cut of the folder with the Module Code given.
