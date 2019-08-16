using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Documents.Templates
{
    class DocumentsSection
    {
        private List<RowSpec> documents;

        public DocumentsSection()
        {
            documents = new List<RowSpec>();
            documents.Add(new RowSpec("", "", "", "", "Документация", 0, ""));
            documents.Add(new RowSpec());
            documents.Add(new RowSpec("", "", "", "", "Сборочный чертеж", 0, ""));
            documents.Add(new RowSpec());
            documents.Add(new RowSpec("", "", "", "", "Схема электрическая принципиальная", 0, ""));
            documents.Add(new RowSpec());
            documents.Add(new RowSpec("А4", "", "", "", "Перечень элементов", 0, ""));
            documents.Add(new RowSpec());
            documents.Add(new RowSpec("", "", "", "", "Ведомость покупных изделий", 0, ""));
        }

        public List<RowSpec> GetDocuments()
        {
            return documents;
        }
    }
}
