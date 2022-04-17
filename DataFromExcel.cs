using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoParser
{
    // КЛАСС УМЕЕТ: ИСПОЛЗУЯ КЛАСС OpenExcel брать данные одной строки и выводить в таблицу
    public class DataFromExcel
    {
        public DataFromExcel() { }
        public DataFromExcel(string Id, string NameUBI , string Description, string Source, string ObjectOfInfluence, string PrivacyViolation, string IntegrityViolation, string AccessibilityViolation)
        {
            this.Id = Id;
            this.NameUBI = NameUBI;
            this.Description = Description;
            this.Source = Source;
            this.ObjectOfInfluence = ObjectOfInfluence;
            this.PrivacyViolation = PrivacyViolation;
            this.IntegrityViolation = IntegrityViolation;
            this.AccessibilityViolation = AccessibilityViolation;
        }

        public override string ToString()
        {
            return $"УБИ.{Id}\t{NameUBI}\nОписание: {Description}\nИсточник: {Source}\nОбъект воздействия угрозы: {ObjectOfInfluence}\nНарушение конфиденциальности: {privacyViolation}\nНарушение целостности: {integrityViolation}\nНарушение доступности: {accessibilityViolation}";
        }

        public string Id { get; set; }
        public string NameUBI { get; set; }                                                        // наименование Угрозы
        public string Description { get; set; }                                                    // описание    
        public string Source { get; set; }                                                         // Источник угрозы (характеристика и потенциал нарушителя)
        public string ObjectOfInfluence { get; set; }                                              // Объект воздействия


        private string privacyViolation;
        private string integrityViolation;
        private string accessibilityViolation;
        public string PrivacyViolation 
        {
            get 
            {
                return privacyViolation;
            }
            set
            {
                if (value == "1") privacyViolation = "да";
                else privacyViolation = "нет";
            }
        }
        public string IntegrityViolation
        {
            get 
            {
                return integrityViolation;
            }
            set
            {
                if (value == "1") integrityViolation = "да";
                else integrityViolation = "нет";
            }

        }
        public string AccessibilityViolation 
        {
            get
            {
                return accessibilityViolation;
            }
            set
            {
                if (value == "1") accessibilityViolation = "да";
                else accessibilityViolation = "нет";
            }
        }
    }
}
