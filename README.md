# NAVAutomation
A dll component that allows communication between NAV objects (e.g. parent/subform) by using Automation Event triggers. 

# Description
In Dynamics NAV (Navision), accessing a parent form from a subform has its challenges. This dll component allows a sort of communication between forms even if they are unrelated. One form can fire a one of the three automation events (Event1, Event2 or Event3) of another form. 

# Usage
Step-by-step
* Identify each form on which you want to expose such communication.
* Identify one form to hold the automation variable that acts as "manager" (or "hub") to facilitate communication between objects
* On each form, create a global automation variable with "WithEvents" property = Yes. Create three functions on each form called "Event1()", "Event2()", "Event3()" and call these functions from the events of automation variable. Why is that is because overriding automation variables sometimes disappears codes behind their events so it is safer to keep such code in different functions.  So it would look like this:
CurrFormAutomation::Event1()
Event1();
CurrFormAutomation::Event2()
Event2();
CurrFormAutomation::Event3()
Event3();
 

