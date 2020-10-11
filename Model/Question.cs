using System;
using LiteDB;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quiz.Model
{
    public class Question
    {
        [BsonId]
        public Guid QuestionId { get; set; }
        public int  QuestionNumber { get; set; }
      
        public string QuestionText { get; set; }
        public string AnswerText { get; set; }
        public string AnswerHTML { get; set; }
        public string Participants { get; set; }
        public bool Active { get; set; }
    }
}
