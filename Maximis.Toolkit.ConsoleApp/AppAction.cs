using System;

namespace Maximis.Toolkit.ConsoleApp
{
    public abstract class AppAction
    {
        public AppAction(ConsoleKey menuKey)
        {
            this.MenuKey = menuKey;
        }

        public object Description { get; set; }

        public ConsoleKey MenuKey { get; set; }

        public abstract void PerformAction(string[] args);
    }
}