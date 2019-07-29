using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoginApplication
{
    public class Product
    {
        private int _idProperty;
        private string _nameProperty;
        private string _pwdProperty;
        private int _identityProperty;
        public int id
        {
            get
            {
                return _idProperty;
            }

            set
            {
                _idProperty = value; ;
            }
        }
        public string name
        {
            get
            {
                return _nameProperty;
            }

            set
            {
                _nameProperty = value;
            }
        }


        public string pwd
        {
            get
            {
                return _pwdProperty;
            }

            set
            {
                _pwdProperty = value;
            }
        }

        public int identity
        {
            get
            {
                return _identityProperty;
            }

            set
            {
                _identityProperty = value; ;
            }
        }

    }
}
