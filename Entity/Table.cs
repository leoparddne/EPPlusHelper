using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;

namespace Entity
{
    public class Table
    {
        [Column]
        [Description(description: "ATest")]
        public DateTime A { get; set; }

        [Column]
        [Description(description: "BTest")]
        public int B { get; set; }

        [Column]
        [Description(description: "BTest")]
        public int C { get; set; }

        [Column]
        [Description(description: "DTest")]
        public TestEnum D { get; set; }
    }

    public enum TestEnum
    {
        a,
        b,
    }
}
