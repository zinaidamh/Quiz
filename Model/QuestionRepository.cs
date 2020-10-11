using LiteDB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Quiz.Model
{
    public class QuestionRepository
    {
        private string liteDBPath = @"Quiz.db";

        public QuestionRepository(string databasePath)
        {
            this.liteDBPath = databasePath;
          
        }

        /// <summary>
        /// Returns all Question Types
        /// </summary>
        /// <returns></returns>
        public List<string> GetQuestionTypes()
        {
            var QuestionTypes = new List<string>();
            QuestionTypes.Add("8-11 years");
            QuestionTypes.Add("12-15 years");
            QuestionTypes.Add("adults");
            return QuestionTypes;
        }

        /// <summary>
        /// Returns a filtered list of all Questions for the matching Question Type and DateTime
        /// </summary>
        /// <param name="QuestionType">Question Type</param>
        /// <param name="datetime">DateTime</param>
        /// <returns></returns>
        public IList<Question> Get(int QuestionNumberFrom, int QuestionNumberTo, bool activeOnly)
        {
            var QuestionsToReturn = new List<Question>();
            using (var db = new LiteDatabase(liteDBPath))
            {
                var Questions = db.GetCollection<Question>("Questions");
                IEnumerable<Question> filteredQuestions;

                if (QuestionNumberFrom!=0 && QuestionNumberTo!=0)
                    filteredQuestions = Questions.FindAll().Where(i => i.QuestionNumber >= QuestionNumberFrom &&
                                                            i.QuestionNumber <= QuestionNumberTo);
                else if (QuestionNumberTo==0)
                    filteredQuestions = Questions.FindAll().Where(i => i.QuestionNumber >= QuestionNumberFrom);
                else if (QuestionNumberFrom == 0)
                    filteredQuestions = Questions.FindAll().Where(i => i.QuestionNumber <= QuestionNumberTo);
                else
                    filteredQuestions = Questions.FindAll();

                if (activeOnly)
                {
                    filteredQuestions = filteredQuestions.Where(q => q.Active == true);
                }

                var sortedQuestions = filteredQuestions.OrderBy(i => i.QuestionNumber);

                foreach (Question QuestionItem in sortedQuestions)
                {
                    QuestionsToReturn.Add(QuestionItem);
                }
                return QuestionsToReturn;
                
            }
        }


        public Question GetOne(int QuestionNumber)
        {
            var QuestionsToReturn = new List<Question>();
            using (var db = new LiteDatabase(liteDBPath))
            {
                var Questions = db.GetCollection<Question>("Questions");



                IEnumerable<Question> filteredQuestions= Questions.FindAll().Where(i => i.QuestionNumber.Equals(QuestionNumber));

                if (filteredQuestions.Any())
                    return filteredQuestions.FirstOrDefault();
                else
                    return null;
             
            }
        }


        /// <summary>
        /// Returns a Collection of Question Items
        /// </summary>
        /// <returns></returns>
        public IList<Question> GetAll()
        {
            var QuestionsToReturn = new List<Question>();
            using (var db = new LiteDatabase(liteDBPath))
            {
                var Questions = db.GetCollection<Question>("Questions");
                var results = Questions.FindAll().OrderBy(i=>i.QuestionNumber);
                foreach (Question QuestionItem in results)
                {
                    QuestionsToReturn.Add(QuestionItem);
                }
                return QuestionsToReturn;
            }
        }

        /// <summary>
        /// Save an Question Item
        /// </summary>
        /// <param name="Question">Question Item</param>
        public void Add(Question Question)
        {
            // Open data file (or create if not exits)
            try
            {
                using (var db = new LiteDatabase(liteDBPath))
                {
                    var QuestionCollection = db.GetCollection<Question>("Questions");
                    // Insert a new Question document
                    QuestionCollection.Insert(Question);
                    IndexQuestion(QuestionCollection);
                }
            }
            catch(Exception ex)
            {
                string str = ex.Message;
            }
        }

        /// <summary>
        /// Update an Existing Question Item
        /// </summary>
        /// 
        /// 
        /// <param name="Question">Question Item</param>
        public void Update(Question Question)
        {
            // Open data file (or create if not exits)
            try
            {
                using (var db = new LiteDatabase(liteDBPath))
                {
                    var QuestionCollection = db.GetCollection<Question>("Questions");
                    
                    // Update an existing Question document
                    QuestionCollection.Update(Question);
                    QuestionCollection.DropIndex("AnswerText");
                }
            }
            catch(Exception ex)
            {
                string str = ex.Message;
            }
        }

        /// <summary>
        /// Delete an Question Item by Question ID (GUID)
        /// </summary>
        /// <param name="QuestionId">Question Id(Guid)</param>
        public void Delete(Guid QuestionId)
        {
            using (var db = new LiteDatabase(liteDBPath))
            {
                var Questions = db.GetCollection<Question>("Questions");
                var value = new LiteDB.BsonValue(QuestionId);//id is an int parameter passed in
                Questions.Delete(value);
                //Questions.Delete(i => i.QuestionId == QuestionId);
            }
        }

        public void DeleteAll()
        {
            using (var db = new LiteDatabase(liteDBPath))
            {
                var Questions = db.GetCollection<Question>("Questions");
           
               //Questions.DeleteMany(i => i.QuestionNumber>400);
            }
        }


        /// <summary>
        /// Index Question
        /// </summary>
        /// <param name="QuestionCollection">Question Collection</param>
        private void IndexQuestion(ILiteCollection<Question> QuestionCollection)
        {
            // Index on QuestionId
            QuestionCollection.EnsureIndex(x => x.QuestionId);

          

            // Index on QuestionNumber
           // QuestionCollection.EnsureIndex(x => x.QuestionNumber);
        }
    }
}
