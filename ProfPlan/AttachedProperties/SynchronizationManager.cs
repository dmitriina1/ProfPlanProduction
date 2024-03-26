using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Controls;

namespace ProfPlan.AttachedProperties
{
    internal class SynchronizationManager
    {
        private readonly Selector _multiSelector;
        private TwoListSynchronizer _synchronizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="SynchronizationManager"/> class.
        /// </summary>
        /// <param name="selector">The selector.</param>
        internal SynchronizationManager(Selector selector)
        {
            _multiSelector = selector;
        }

        /// <summary>
        /// Starts synchronizing the list.
        /// </summary>
        public void StartSynchronizingList()
        {
            IList list = MultiSelectorBehaviours.GetSynchronizedSelectedItems(_multiSelector);

            if (list != null)
            {
                _synchronizer = new TwoListSynchronizer(GetSelectedItemsCollection(_multiSelector), list);
                _synchronizer.StartSynchronizing();
            }
        }

        /// <summary>
        /// Stops synchronizing the list.
        /// </summary>
        public void StopSynchronizing()
        {
            _synchronizer.StopSynchronizing();
        }

        public static IList GetSelectedItemsCollection(Selector selector)
        {
            if (selector is MultiSelector)
            {
                return (selector as MultiSelector).SelectedItems;
            }
            else if (selector is ListBox)
            {
                return (selector as ListBox).SelectedItems;
            }
            else
            {
                throw new InvalidOperationException("Target object has no SelectedItems property to bind.");
            }
        }
    }
}
