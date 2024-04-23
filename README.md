# STE_WorkflowAnalyzerRules
## Custom workflow analyzer rules for UiPath Studio
### Dependencies
1. UiPath.Activities.API.Base
2. UiPath.API.Base
3. UiPath.Robot.Activities.API
4. UiPath.Studio.Activities.API
5. UiPath.Studio.API.Base
   
### Rules to enforce
1. Project naming convention
2. Variable/argument naming convention
3. Proper variable/argument types (No generic or object values)
4. Activity renamed from default name
5. Annotations in all workflows
6. Log messages used in all workflows
7. Project using REFramework rule
8. Passwords are all stored as SecureString types
9. Input dialog taking Password inputs take in masked inputs
10. No hardcoded delay activities
11. Excel application usage are all set to execute in background
12. Outlook email are processed FIFO
13. Outlook emails are retrived using defined filters
14. Sensitive information are not stored in log messages
