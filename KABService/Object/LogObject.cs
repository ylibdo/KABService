using System;

namespace UtilityLibrary.Log
{
    class LogObject
    {
        public enum LogType { Information, Warning, Error, Exception }
        public LogType Type { get; set; }
        public string Message { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public string CreatedBy { get; set; }

        public LogObject(LogType _logType, string _message, string _createdBy = "")
        {
            Type = _logType;
            Message = _message;
            CreatedDateTime = DateTime.Now;
            CreatedBy = String.IsNullOrEmpty(_createdBy) ? "System" : _createdBy ;
        }

        public string ToString(string _identifier)
        {
            string returnObject = String.Concat(
                "Source: ", _identifier, Environment.NewLine,
                "Log Type: ", this.Type.ToString(), Environment.NewLine,
                "Message: ", this.Message, Environment.NewLine,
                "CreatedBy: ", this.CreatedBy, " ,Timestamp: ", this.CreatedDateTime.ToString()
                );
            return returnObject;
        }
    }
}
