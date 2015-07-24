using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;


namespace NAVAutomationHandler
{
    /*
     * Turn "WithEvents" property on in order to see available events. 
     * In case you want to add more events: 
     *   1 - add them to UserEvents first with proper Display IDs
     *   2 - add definitions to $safeprojectname$ class header
     *   3 - update TotalEventCount constant
     *   4 - update FireEvent function
     */
    public delegate void EventDel();
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIDispatch)]
    public interface UserEvents
    {
        [DispId(1001)]
        void Event1();
        [DispId(1002)]
        void Event2();
        [DispId(1003)]
        void Event3();
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface INAVAutomation
    {
        /*
         * AddAutomation should be called on manager object.
         * The manager object maintains a list-of same-type objects idenfied by their unique name. 
         * If curent automation object already has a name, the second argument (AutomatName2) is skipped.         
         */
        [DispId(10)]
        void AddAutomation(INAVAutomation AutomationObj2, string AutomatName2);

        [DispId(15)]
        Boolean HasAutomation(string AutomatName2);

        /*
         * Call this on manager object to remove an automation object from the array-list. 
         */
        [DispId(20)]
        void RemoveAutomation(string AutomatName2);     

        /* 
         * Removes all automation objects from the array within current object. 
         * It does not clear hashtables implicitly            
         */
        [DispId(22)]
        void RemoveAutomations();

        /* 
         * You can call this if you iterate on automation objects on current manager object.         
         */ 
        [DispId(23)]
        int GetAutomationCount();

        
        /*
         * Returns automation object at specified index. You can use it together with GetAutomationCount
         * to iterate over automtion list. Keep in mind that the index here starts from "1": NAV Style. 
         */
        [DispId(24)]
        INAVAutomation GetAutomationAtIndex(int Index2);
        
        /*
         * This can be called on manager object. It finds automation object by name from the array list. 
         * e.g. : mainAutomat.GetAutomation("commentlinesubform");
         */
        [DispId(26)]
        INAVAutomation GetAutomation(string AutomatName2);


        /*
         * Sets name for current automation object. 
         * This function overrides the name that is set by AddAutomation function       
         */ 
        [DispId(30)]
        void SetAutomationName(string AutomatName2);

        [DispId(32)]
        string GetAutomationName();

        /*
         * Each automation object maintains its own variable in a hashtable. 
         * e.g. CurrFormAutomat.AddKey("CustomerNo",cust."No.");
         */  
        [DispId(40)]
        void AddKey(string Key2, string Value2);

        /*
         * This variation of AddKey allows you to set variable on a different automation objects.  
         * You can call this from the manager automation. The manager aut. will scan object-list by AutomationName2
         * e.g. mainform.AddKey_2("commentform","highlight","true"); 
         */
        [DispId(42)]
        void AddKey(string AutomationName2, string Key2, string Value2);


        /*
         * Removes a key value from the current automation object hashtable.          
         */
        [DispId(50)]
        void RemoveKey(string Key2);

        /*
         * A variation of RemoveKey to be called from the manager object. 
         * It finds the given automation obj. by name and remove given key from the hashtable
         */ 
        [DispId(52)]
        void RemoveKey(string AutomationName2, string Key2);

        /*
         * Clears the hashtable of current object 
         */ 
        [DispId(56)]
        void EmptyHashtable();

        /*
         * 
         * This can be called on manager object. It finds the automation by name and clears the hashtable
         * e.g. mainAutomat.EmptyHashtable("salesform");
         */
        [DispId(58)]
        void EmptyHashtable(string AutomationName2);

        /*
         * Finds each automation object within array-list and clears each hash tables. 
         * Use this when you need to reset all variables but you want to keep the automation objects
         */        
        [DispId(59)]
        void EmptyHashtables();

        /*
         * check if hashtable on current automation object has a specified key value. 
         */ 
        [DispId(60)]
        Boolean HasKey(string Key2);

        /*
         * This be called on the manager object. 
         * It finds the given automation by name and checks if its hashtable has specified key
         */ 
        [DispId(62)]
        Boolean HasKey(string AutomationName2, string Key2);

        /*
         * Retrieves value from hashtable on current automation variable
         */ 
        [DispId(70)]
        string GetValue(string Key2);

        /*
         * To be called on manager object. It finds the given automation by name and retrieves the key value. 
         */ 
        [DispId(72)]
        string GetValue(string AutomationName2, string Key2);

        /*
         * Returns current automation, hashtable length. 
         */ 
        [DispId(90)]
        int GetHashLength();

        /*
         * This can be called from managet object.  
         * It finds the given automation object and returns the hashtable length 
         */ 
        [DispId(92)]
        int GetHashLength(string AutomationName2);

        /*
         * Retrieves hashtable key by given index. Even though this is a hashtable, you can use this function
         * in order to iterate keys from hashtable. This is a costly read: O(n)
         * Index starts from 1. 
        */
        [DispId(100)]
        string GetKeyAtIndex(int Index2);

        /*
         * This can be called on manager object. 
         * It finds the given automation object and retrieves the "key" at given index. 
         * Index starts from 1
         */
        [DispId(102)]
        string GetKeyAtIndex(string AutomationName2, int Index2);  

        /*
         * This calls given event index on current automation object. 
         * e.g. CurrFormAutomat.FireEvent(2); and it will call Event2();
         */ 
        [DispId(600)]
        void FireEvent(int EventID2);

        /*
         * This can be called on the manager object. 
         * It finds the automation object in arraylist and fires the event with given id. 
         * e.g. mainform.FireEvent_2("historyform",2);
         */        
        [DispId(602)]
        void FireEvent(string AutomationName2, int EventID2);

    }


    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("NAVFormHandler.NAVFormHandler")]
    [ComSourceInterfaces(typeof(UserEvents))]

    public class $safeprojectname$ : INAVAutomation
    {
        public event EventDel event1;
        public event EventDel event2;
        public event EventDel event3;
        public const int TotalEventCount = 3;

        //global variables
        private string AutomationName = null;
        private ArrayList AutomationList = null;
        private Hashtable Hash = null;

        public $safeprojectname$()
        {
            this.AutomationList = new ArrayList();
            this.Hash = new Hashtable();
            this.AutomationName = "";
        }

        public void SetAutomationName(string AutomationName2)
        {
            this.AutomationName = AutomationName2;
        }

        public string GetAutomationName()
        {
            return this.AutomationName;
        }


        public void AddAutomation(INAVAutomation Automation2, string AutomationName2)
        {
            if ((Automation2 != null))
            {
                if ((AutomationName2 != "") && (Automation2.GetAutomationName() == ""))
                    Automation2.SetAutomationName(AutomationName2);

                //objects must have a name in order to be added to the automation array
                if ((Automation2.GetAutomationName() == ""))
                    throw new Exception("AddAutomation failed. Object must have a name");                  
                
                //each automation must have a unique name
                if (!AutomationList.Contains(Automation2))          
                    AutomationList.Add(Automation2);               
                else
                {
                    //Possible exception
                    //throw new Exception("Automation " + AutomationName2 +" exists.");
                }
            }
            else            
                throw new NullReferenceException("AddAutomation failed due to null reference.");
            
        }
        public Boolean HasAutomation(string AutomationName2)
        {
            if (AutomationName2 == "")
                return true;

            $safeprojectname$ Aut2 = new $safeprojectname$();
            Aut2.SetAutomationName(AutomationName2);
            return AutomationList.Contains(Aut2);                                  
        }



        public INAVAutomation GetAutomation(string AutomationName2)
        {
            if (AutomationName2 == "")
                return null;

            $safeprojectname$ aut = new $safeprojectname$();
            aut.SetAutomationName(AutomationName2);

            int index = AutomationList.IndexOf(aut);            
            if (index >= 0)            
                return (INAVAutomation)AutomationList[index];
            
            return null;
        }

        //NAV style, index starts at 1. 
        public INAVAutomation GetAutomationAtIndex(int Index2)
        {
            return (INAVAutomation)AutomationList[Index2 - 1];
        }

        public void RemoveAutomation(string AutomationName2)
        {
            if (AutomationName2 == "")
                return;
            INAVAutomation obj2 = (INAVAutomation) GetAutomation(AutomationName2);
            if (obj2 != null)
                AutomationList.Remove(obj2);
        }

        public void RemoveAutomations()
        {
            AutomationList.Clear();
        }

        public int GetAutomationCount()
        {
            return AutomationList.Count;
        }


        public void AddKey(string Key2, string Value2)
        {
            Hash[Key2] = Value2;
        }

        public void AddKey(string AutomationName2, string Key2, string Value2)
        {
            INAVAutomation obj2 = this.GetAutomation(AutomationName2);
            if (obj2 != null)
                obj2.AddKey(Key2, Value2);
        }

        public string GetValue(string Key2)
        {
            if (Hash.ContainsKey(Key2))            
                return (String)Hash[Key2];
            return "";
        }

        public string GetValue(string AutomationName2, string Key2)
        {
            INAVAutomation obj2 = this.GetAutomation(AutomationName2);
            if (obj2 != null)
                return obj2.GetValue(Key2);
            return "";
        }

        public void RemoveKey(string Key2)
        {
            Hash.Remove(Key2);
        }

        public void RemoveKey(string AutomationName2, string Key2)
        {
            INAVAutomation obj2 = this.GetAutomation(AutomationName2);
            if (obj2 != null)
                obj2.RemoveKey(Key2);
          
        }
        public void EmptyHashtable()
        {
            Hash.Clear();
        }

        public void EmptyHashtable(string AutomationName2)
        {
            INAVAutomation obj2 =  GetAutomation(AutomationName2);
            if (obj2 != null)
                obj2.EmptyHashtable();
        }

        public void EmptyHashtables()
        {
            foreach (INAVAutomation obj2 in this.AutomationList)
            {
                obj2.EmptyHashtable();
            }
        }



        public bool HasKey(string Key2)
        {
            return Hash.Contains(Key2);
        }

        public bool HasKey(string AutomationName2, string Key2)
        {
            if (AutomationName2 != "")
            {
                INAVAutomation form2 = this.GetAutomation(AutomationName2);
                if (form2 != null)               
                    return form2.HasKey(Key2);
                
            }
            return false;
        }

        public string toString()
        {
            //used in comparison in arraylist when adding/removing objects to the list
            return this.AutomationName;
        }


        public void FireEvent(string AutomationName2, int EventID2)
        {
            if (AutomationName2 == "")            
                return;            

            if ((EventID2 <= 0) || (EventID2 > TotalEventCount))            
                return;

            INAVAutomation obj2 = GetAutomation(AutomationName2);
            if (obj2 != null)
                obj2.FireEvent(EventID2);            
        }

        public void FireEvent(int EventID2)
        {
            if (EventID2 == 1)
                event1();
            else if (EventID2 == 2)
                event2();
            else if (EventID2 == 3)
                event3();
        }

        public override bool Equals(object obj2)
        {
           
           INAVAutomation aut2 = (INAVAutomation)obj2;
           return aut2.GetAutomationName() == this.GetAutomationName();
            
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }


        //You might use it to iterate on hash table from "1" to "hash.count"
        public int GetHashLength()
        {
            return Hash.Count;
        }

        public int GetHashLength(string AutomationName2)
        {
            INAVAutomation obj2 = this.GetAutomation(AutomationName2);
            if (obj2 != null)            
                return obj2.GetHashLength();            
            return 0;
        }


        public string GetKeyAtIndex(int Index2)
        {
            //this is a costly method but you might want to iterate over hashtable values 
            //at some moment.. 
            if ((Index2 <= 0) || (Index2 > Hash.Count))
                return "";
            int counter2 = 0;
            foreach (string keyValue in Hash.Keys)
            {
                counter2 += 1;
                if (counter2 == Index2)
                    return keyValue;
            }
            return "";
        }

        public string GetKeyAtIndex(string AutomationName2, int Index2)
        {
            INAVAutomation obj2 = this.GetAutomation(AutomationName2);
            if (obj2 != null)            
                return obj2.GetKeyAtIndex(Index2);               
            return "";
        }


    }
}
