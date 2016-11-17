using System.Collections.Generic;

namespace ApiExamples.TestData
{
    public class TestClass5
    {
        public TestClass5(string name, int age, IEnumerable<TestClass5> children)
        {
            this.Name = name;
            this.Age = age;
            this.Children = children;
        }

        public string Name { get; set; }

        public int Age { get; set; }

        public IEnumerable<TestClass5> Children { get; set; }
    }
}
