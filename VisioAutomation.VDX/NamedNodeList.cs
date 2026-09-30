using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.VDX
{
    public class NamedNodeList<T> : Elements.Node where T : Elements.Node
    {
        private readonly List<T> items = new List<T>();
        private readonly System.Collections.ObjectModel.ReadOnlyCollection<T> view;
        private readonly System.Func<T, string> func_get_name;
        private static readonly System.StringComparer comparer = System.StringComparer.InvariantCultureIgnoreCase;

        public NamedNodeList(System.Func<T, string> func_get_name)
        {
            if (func_get_name == null)
            {
                throw new System.ArgumentNullException(nameof(func_get_name));
            }
            this.func_get_name = func_get_name;
            this.view = this.items.AsReadOnly();
        }

        public bool ContainsName(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }
            // Names are mutable, so the items themselves are the source of truth.
            return this.items.Any(item => comparer.Equals(this.func_get_name(item), name));
        }

        public bool Contains(T item)
        {
            if (item == null)
            {
                throw new System.ArgumentNullException(nameof(item));
            }
            return item._parent == this;
        }

        protected void ValidateItem(T item)
        {
            if (item == null)
            {
                throw new System.ArgumentNullException(nameof(item));
            }
            if (item._parent != null)
            {
                throw new System.ArgumentException("Item already belongs to a collection", nameof(item));
            }
        }

        protected void ValidateName(string name)
        {
            if (this.ContainsName(name))
            {
                throw new System.ArgumentException("Already contains an item with that name", nameof(name));
            }
        }

        public virtual void Add(T item)
        {
            this.ValidateItem(item);
            this.ValidateName(this.func_get_name(item));
            this.items.Add(item);
            item._parent = this;
        }

        public T this[string name]
        {
            get
            {
                if (name == null)
                {
                    throw new System.ArgumentNullException(nameof(name));
                }
                var item = this.items.FirstOrDefault(candidate => comparer.Equals(this.func_get_name(candidate), name));
                if (item == null)
                {
                    throw new KeyNotFoundException("No item with that name");
                }
                return item;
            }
        }

        public int Count => this.items.Count;

        public IEnumerable<T> Items => this.view;
    }
}
