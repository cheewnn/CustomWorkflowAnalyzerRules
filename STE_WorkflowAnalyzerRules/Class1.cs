using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Deployment.Internal;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using UiPath.Studio.Activities.Api;
using UiPath.Studio.Activities.Api.Analyzer;
using UiPath.Studio.Activities.Api.Analyzer.Rules;
using UiPath.Studio.Analyzer.Models;
using Humanizer;
using System.Reflection;
using UiPath.Robot.Activities.Api;

namespace STE_WorkflowAnalyzerRules
{
    public class STE_RuleRepository : IRegisterAnalyzerConfiguration
    {

        // 00. Initialize workflowAnalyzerConfigService
        public void Initialize(IAnalyzerConfigurationService workflowAnalyzerConfigService)
        {
            if (!workflowAnalyzerConfigService.HasFeature("WorkflowAnalyzerV6"))
                return;

            #region 00.01. ProjectNamingRule "Project Naming Convention" "STE-NMG-001"
            var ProjectNamingRule = new Rule<IProjectModel>("01. Project Naming Convention", "STE-NMG-001", InspectProjectName);
            ProjectNamingRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            ProjectNamingRule.RecommendationMessage = "Ensure that your project naming follows the following format DepartmentName_MeaningfulDescription" + System.Environment.NewLine +
                                                      "You can change your project name in Project Panel > Project Settings > General > Name";
            ProjectNamingRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            ProjectNamingRule.Parameters.Add("Project_Name_Regex", new Parameter()
            {
                DefaultValue = "^([A-Za-z0-9]+)_([A-Za-z0-9]+)$",
                Key = "Project_Name_Regex",
                LocalizedDisplayName = "Proper project name format"
            });
            ProjectNamingRule.Parameters.Add("List_Of_Departments", new Parameter()
            {
                DefaultValue = "FIN|HR|IT",
                Key = "List_Of_Departments",
                LocalizedDisplayName = "List of departments in your organization, delimit list with '|' character."
            });
            workflowAnalyzerConfigService.AddRule<IProjectModel>(ProjectNamingRule);
            #endregion

            #region 00.02.1. VariableNamingRule "Variable Naming Convention" "STE-NMG-002"
            var VariableNamingRule = new Rule<IActivityModel>("02.1. Variable Naming Convention", "STE-NMG-002", InspectVariableName);
            VariableNamingRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            VariableNamingRule.RecommendationMessage = "Ensure that the variable name follows this format: VariableType_MeaningfulDescription";
            VariableNamingRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            VariableNamingRule.Parameters.Add("Variable_Name_Regex", new Parameter()
            {
                DefaultValue = "^([A-Za-z0-9]+)_([A-Za-z0-9]+)$",
                Key = "Variable_Name_Regex",
                LocalizedDisplayName = "Proper variable name format"
            });
            VariableNamingRule.Parameters.Add("List_Of_VarTypes", new Parameter()
            {
                DefaultValue = "dt|list|dict|str|int|dbl|arr|bool|dr|date|dec",
                Key = "List_Of_Variables",
                LocalizedDisplayName = "List of valid variables, delimit list with '|' character."
            });
            workflowAnalyzerConfigService.AddRule<IActivityModel>(VariableNamingRule);
            #endregion

            #region 00.02.2. ArgumentNamingRule "Argument not following proper naming convention" "STE-NMG-003"
            var ArgumentNamingRule = new Rule<IWorkflowModel>("02.2. Argument not following proper naming convention", "STE-NMG-003", InspectArgumentName);
            ArgumentNamingRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            ArgumentNamingRule.RecommendationMessage = "Ensure that argument naming starts with argument direction: in, out or io";
            ArgumentNamingRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            ArgumentNamingRule.Parameters.Add("Argument_Name_Starting", new Parameter()
            {
                DefaultValue = "in_|out_|io_",
                Key = "Argument_Name_Starting",
                LocalizedDisplayName = "Valid argument start names, delimit list with '|' character."
            });
            workflowAnalyzerConfigService.AddRule(ArgumentNamingRule);

            #endregion

            #region 00.03.1. ProperVariableTypeRule "03.1. No variables with Object/GenericValue types" "STE-USG-001"
            var ProperVariableTypeRule = new Rule<IActivityModel>("03.1. Variables with Object/GenericValue types", "STE-USG-001", InspectProperVariableType);
            ProperVariableTypeRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            ProperVariableTypeRule.RecommendationMessage = "It is recommended to specify explicit variable types to enhance workflow clarity and maintainability." +Environment.NewLine+
                                                           "Using specific types provides better validation and helps prevent unexpected errors during execution." + Environment.NewLine+
                                                           "Please review the variable/argument type and update it to the most appropriate data type for your use case.";
            ProperVariableTypeRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            ProperVariableTypeRule.Parameters.Add("Invalid_Types", new Parameter()
            {
                DefaultValue = "Object|GenericValue",
                Key = "Invalid_Types",
                LocalizedDisplayName = "Invalid variable or arguments data types, delimit list with '|' character."
            });
            workflowAnalyzerConfigService.AddRule(ProperVariableTypeRule);
            #endregion

            #region 00.03.2. ProperVariableTypeRule "03.2. No variables with Object/GenericValue types" "STE-USG-002"
            var ProperArgumentTypeRule = new Rule<IWorkflowModel>("03.2. Arguments with Object/GenericValue types", "STE-USG-002", InspectProperArgumentType);
            ProperArgumentTypeRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            ProperArgumentTypeRule.RecommendationMessage = "It is recommended to specify explicit argument types to enhance workflow clarity and maintainability." +Environment.NewLine+
                                                           "Using specific types provides better validation and helps prevent unexpected errors during execution." +Environment.NewLine+
                                                           "Please review the argument type and update it to the most appropriate data type for your use case.";
            ProperArgumentTypeRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            ProperArgumentTypeRule.Parameters.Add("Invalid_Types", new Parameter()
            {
                DefaultValue = "Object|GenericValue",
                Key = "Invalid_Types",
                LocalizedDisplayName = "Invalid variable or arguments data types, delimit list with '|' character."
            });
            workflowAnalyzerConfigService.AddRule(ProperArgumentTypeRule);
            #endregion

            #region 00.04.1. RenamedActivityNameRule "Activities are renamed from default name" "STE-MRD-001"
            var RenamedActivityNameRule = new Rule<IActivityModel>("04.1. Activities are renamed from default names", "STE-MRD-001", InspectActivityRenamed);
            RenamedActivityNameRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            RenamedActivityNameRule.RecommendationMessage = "It is important to use meaningful and descriptive names for your activities to enhance workflow readability and maintainability." +Environment.NewLine+
                                                            "Using default display names may lead to confusion and make it harder to understand the purpose of each activity.";
            RenamedActivityNameRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IActivityModel>(RenamedActivityNameRule);
            #endregion

            #region 00.04.2. AnnotationsInWorkflowRule "Annotations in activities must be present in workflows" "STE-MRD-002"
            var AnnotationsInWorkflowRule = new Rule<IWorkflowModel>("04.2. Annotations in activities must be present in workflows", "STE-MRD-002", InspectAnnotationsInWorkflow);
            AnnotationsInWorkflowRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            AnnotationsInWorkflowRule.RecommendationMessage = "Annotations are essential for documenting your workflow, providing context, and making it more understandable for both you and your team." +Environment.NewLine+
                                                              "Consider adding additional annotations to explain key steps, decisions, or any other important information.";
            AnnotationsInWorkflowRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            AnnotationsInWorkflowRule.Parameters.Add("Minimum_Annotations_Count", new Parameter()
            {
                DefaultValue = "2",
                Key = "Minimum_Annotations_Count",
                LocalizedDisplayName = "Minimum number of annotated activities per workflow (Integer)"
            });
            workflowAnalyzerConfigService.AddRule<IWorkflowModel>(AnnotationsInWorkflowRule);
            #endregion

            #region 00.04.3. LogMsgInWorkflowRule "Log Messages must be present in workflows" "STE-DCP-001"
            var LogMsgInWorkflowRule = new Rule<IWorkflowModel>("04.3. Log Messages must be present in workflows", "STE-DBP-001", InspectLogsInWorkflow);
            LogMsgInWorkflowRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            LogMsgInWorkflowRule.RecommendationMessage = "Logging is crucial for capturing important information during workflow execution." + Environment.NewLine +
                                                         "Consider adding additional log messages to record key events, decisions, or any other significant details.";
            LogMsgInWorkflowRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            LogMsgInWorkflowRule.Parameters.Add("Minimum_Logs_Count", new Parameter()
            {
                DefaultValue = "2",
                Key = "Minimum_Logs_Count",
                LocalizedDisplayName = "Minimum number of Log Message activites per workflow (Integer)"
            });
            workflowAnalyzerConfigService.AddRule<IWorkflowModel>(LogMsgInWorkflowRule);
            #endregion

            #region 00.05. UsingREFRule "Project is using REFramework" "STE-DBP-002"
            var UsingREFRule = new Rule<IProjectModel>("05. Project is using REFramework", "STE-DBP-002", InspectFrameworkType);
            UsingREFRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            UsingREFRule.RecommendationMessage = "The Robotic Enterprise Framework (REFramework) is a powerful template in UiPath designed for building scalable and maintainable automation solutions." +Environment.NewLine+
                                                 "It enforces best practices, promotes code reusability, and enhances error handling.";
            UsingREFRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IProjectModel>(UsingREFRule);
            #endregion

            #region 00.06.1. PasswordSecureStringRule "Password variables are of 'SecureString' type." "STE-SEC-001"
            var PasswordSecureStringRule = new Rule<IActivityModel>("06.1. Password variables are of 'SecureString' type.", "STE-SEC-001", InspectPasswordVariables);
            PasswordSecureStringRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Error;
            PasswordSecureStringRule.RecommendationMessage = "For security reasons, it is mandatory to use the 'SecureString' data type for storing password variables." +Environment.NewLine+
                                                             "Using 'SecureString' helps protect sensitive information by encrypting it in memory and providing additional security against potential vulnerabilities.";
            PasswordSecureStringRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            PasswordSecureStringRule.Parameters.Add("Password_Variable_Contains", new Parameter()
            {
                DefaultValue = "password|pw|pass",
                Key = "Password_Variable_Contains",
                LocalizedDisplayName = "Words that are present in a password variable, delimited by '|' char"
            });
            workflowAnalyzerConfigService.AddRule<IActivityModel>(PasswordSecureStringRule);
            #endregion

            #region 00.06.2. InputDialogIsPasswordRule "If user input dialog is for a password credential, 'IsPassword' property is set to 'true' and output 'Result' argument is set to a 'SecureString' variable" "STE-SEC-002"
            var InputDialogIsPasswordRule = new Rule<IActivityModel>("06.2. If user input dialog is for a password credential, 'IsPassword' property is set to 'true' and output 'Result' argument is set to a 'SecureString' variable", "STE-SEC-002", InspectPasswordInputDialogVariables);
            InputDialogIsPasswordRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Error;
            InputDialogIsPasswordRule.RecommendationMessage = "Ensuring secure handling of passwords is crucial for protecting sensitive information." +Environment.NewLine+
                                                              "When using an Input Dialog for password input, follow best practices by setting the 'IsPassword' property to true and ensuring that the output result variable is of type 'SecureString'.";
            InputDialogIsPasswordRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IActivityModel>(InputDialogIsPasswordRule);

            #endregion

            #region 00.07. NoDelayActivityRule "No delay activities should be used" "STE-BDP-003"
            var NoDelayActivityRule = new Rule<IActivityModel>("07. No delay activities should be used", "STE-DBP-003", InspectNoDelayActivity);
            NoDelayActivityRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            NoDelayActivityRule.RecommendationMessage = "Using Delay activities can introduce unnecessary delays and make your workflows less efficient." +Environment.NewLine+
                                                        "Instead, leverage activities that are specifically designed for waiting, such as Check App State or Element Exists.";
            NoDelayActivityRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IActivityModel>(NoDelayActivityRule);
            #endregion

            #region 00.08. ExcelVisibleRule "Excel applications are always run in the background" "STE-BDP-004"
            var ExcelVisibleRule = new Rule<IActivityModel>("08. Excel applications are always run in the background", "STE-DBP-004", InspectExcelVisible);
            ExcelVisibleRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            ExcelVisibleRule.RecommendationMessage = "It is recommended to use the Excel Application Scope activity when working with Excel files to ensure better performance and reliability." +Environment.NewLine+
                                                     "Additionally, set the 'Show Excel Window'/'Visible' property to 'False' within the Excel Application Scope to run Excel operations in the background without displaying the Excel application window.";
            ExcelVisibleRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IActivityModel>(ExcelVisibleRule);
            #endregion

            #region 00.09. OutlookFifoRule "Outlook: Emails are processed in a FIFO manner" "STE-BDP-005"
            var OutlookFifoRule = new Rule<IActivityModel>("09. Outlook: Emails are processed in a FIFO manner", "STE-DBP-005", InspectOutlookGetMail);
            OutlookFifoRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            OutlookFifoRule.RecommendationMessage = "To ensure that Outlook mails are processed in a First-In-First-Out (FIFO) manner, it is recommended to set the OrderByDate property of the Get Outlook Mail Messages activity to 'OldestFirst'." +Environment.NewLine+
                                                    "This ensures that the oldest emails are processed first.";
            OutlookFifoRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IActivityModel>(OutlookFifoRule);
            #endregion

            #region 00.10. OutlookFiltersRule "Outlook: Filters are used when querying emails from Outlook" "STE-BDP-006"
            var OutlookFiltersRule = new Rule<IActivityModel>("10. Outlook: Filters are used when querying emails from Outlook", "STE-DBP-006", InspectOutlookGetMailFilters);
            OutlookFiltersRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Warning;
            OutlookFiltersRule.RecommendationMessage = "It is recommended to apply a filter when querying mail messages using the Get Outlook Mail Messages activity." +Environment.NewLine+
                                                       "This allows you to narrow down the search and retrieve only the relevant emails, improving performance and ensuring that your automation processes the required data.";
            OutlookFiltersRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            workflowAnalyzerConfigService.AddRule<IActivityModel>(OutlookFiltersRule);
            #endregion

            #region 00.11. SensitiveNotLogRule "Sensitive information handled by the bot is not written out to external files unnecessary" "STE-SEC-003"
            var SensitiveNotLogRule = new Rule<IActivityModel>("11. Sensitive information handled by the bot is not written out to external files unnecessary", "STE-SEC-003", InspectSensitiveNotLogged);
            SensitiveNotLogRule.DefaultErrorLevel = System.Diagnostics.TraceLevel.Error;
            SensitiveNotLogRule.RecommendationMessage = "It is crucial to avoid logging or displaying sensitive information using Log Message or Write Line activities, as this can pose a security risk." +Environment.NewLine+
                                                        "Ensure that your automation processes sensitive data securely and follows best practices for handling confidential information.";
            SensitiveNotLogRule.ApplicableScopes = new List<string> { RuleConstants.BusinessRule };
            SensitiveNotLogRule.Parameters.Add("Variable_Names", new Parameter()
            {
                DefaultValue = "NRIC|Email|Phone|Mobile",
                Key = "Variable_Names",
                LocalizedDisplayName = "Variable names"
            });
            workflowAnalyzerConfigService.AddRule<IActivityModel>(SensitiveNotLogRule);

            #endregion

        }

        #region 01. START OF InspectProjectName for ProjectNamingRule
        private InspectionResult InspectProjectName(IProjectModel projectToInspect, Rule configuredRule)
        {
            var configProjectNameFormat = configuredRule.Parameters["Project_Name_Regex"]?.Value;
            var configListOfDepartments = configuredRule.Parameters["List_Of_Departments"]?.Value;
            var lst_Messages = new List<InspectionMessage>();
            string projectName = projectToInspect.DisplayName;


            // If Project_Name_Regex param is not set, then set error to false
            if (string.IsNullOrWhiteSpace(configProjectNameFormat))
            {
                return new InspectionResult() { HasErrors = false };
            }


            //Check if project name fits the regex criteria
            //configured regex is ^([A-Za-z0-9]+)_([A-Za-z0-9]+)$
            Match matchProjectName = Regex.Match(projectName, configProjectNameFormat);
            if (matchProjectName.Success)
            {
                Console.WriteLine("Project name has a prefix and suffix");
            }
            else
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = "Project name does not follow this format: {dept/team name}_{Meaningful description}"
                });
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "Ensure that your project naming follows the following format DepartmentName_MeaningfulDescription" + System.Environment.NewLine +
                                            "You can change your project name in Project Panel > Project Settings > General > Name",
                    ErrorLevel = configuredRule.ErrorLevel
                };
            }

            //Check if department name extracted is within the list of departments
            string deptName = matchProjectName.Groups[1].Value;
            System.Console.WriteLine(deptName);

            if (!string.IsNullOrWhiteSpace(configListOfDepartments))
            {
                List<string> lst_Departments = configListOfDepartments.Split('|').Select(
                    dept => dept.Trim()
                ).ToList();
                if (lst_Departments.Contains(deptName)) //Department name set in project name is in list of departments param
                {
                    System.Console.WriteLine($"{deptName} is a valid department name");
                    return new InspectionResult() { HasErrors = false };
                }
                else //Department name is invalid, return a recommendation message
                {
                    lst_Messages.Add(new InspectionMessage()
                    {
                        Message = "Department name in project name is invalid."
                    });
                    return new InspectionResult()
                    {
                        HasErrors = true,
                        InspectionMessages = lst_Messages,
                        RecommendationMessage = "Ensure that your project naming follows the following format DepartmentName_MeaningfulDescription" + System.Environment.NewLine +
                                                "You can change your project name in Project Panel > Project Settings > General > Name"+ System.Environment.NewLine +
                                                $"List of available department names: {configListOfDepartments.Replace("|",", ")}",
                        ErrorLevel = configuredRule.ErrorLevel
                    };
                }
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false
                };
            }
        }
        #endregion

        #region 02.1 START OF InspectVariableName for VariableNamingRule
        private InspectionResult InspectVariableName(IActivityModel activityToInspect, Rule configuredRule)
        {
            var configVariableNameRegex = configuredRule.Parameters["Variable_Name_Regex"]?.Value;
            var configVariableTypes = configuredRule.Parameters["List_Of_VarTypes"]?.Value;
            var lst_Messages = new List<InspectionMessage>();

            if (String.IsNullOrWhiteSpace(configVariableNameRegex)) //if regex not configured
            {
                return new InspectionResult() { HasErrors = false };
            }


            foreach (var variable in activityToInspect.Variables)
            {
                Match matchVariableName = Regex.Match(variable.DisplayName, configVariableNameRegex);
                if (!matchVariableName.Success)
                {
                    lst_Messages.Add(new InspectionMessage()
                    {
                        Message = "Variable name '" + variable.DisplayName + "' does not follow this format: VariableType_MeaningfulDescription"
                    });
                }
                else
                {
                    if (String.IsNullOrWhiteSpace(configVariableTypes)) // if list of variable types is not configured
                    {
                        return new InspectionResult() { HasErrors = false };
                    }

                    // Check that variable type assigned is one of the predefined variable types set in configVariableTypes
                    string varType = matchVariableName.Groups[1].Value;
                    List<string> lst_Types = configVariableTypes.Split('|').Select(
                        dept => dept.Trim()
                    ).ToList();
                    if (!lst_Types.Contains(varType)) //Variable type is invalid, return a recommendation message
                    {
                        lst_Messages.Add(new InspectionMessage()
                        {
                            Message = $"Variable type '{varType}' written in variable name '{variable.DisplayName}' is invalid."
                        });
                    }
                }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "Ensure that the variable name follows this format: VariableType_MeaningfulDescription" + System.Environment.NewLine + 
                                            $"list of valid Variable Types: {configVariableTypes.Replace("|", ",")}",
                    ErrorLevel = configuredRule.ErrorLevel
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }



        }
        #endregion

        #region 02.2 START OF InspectArgumentName for ArgumentNamingRule
        private InspectionResult InspectArgumentName(IWorkflowModel workflowToInspect, Rule configuredRule)
        {
            var configArgumentStartingValues = configuredRule.Parameters["Argument_Name_Starting"]?.Value;
            var lst_Messages = new List<InspectionMessage>();
            if (string.IsNullOrWhiteSpace(configArgumentStartingValues))
            {
                return new InspectionResult { HasErrors = false };
            }
            List<string> lst_ArgumentStartingValues = configArgumentStartingValues.Split('|').ToList();
            foreach (var argument in workflowToInspect.Arguments)
            {
                bool argumentValid = false;
                foreach (var ArgumentStartingValue in lst_ArgumentStartingValues)
                {
                    if (argument.DisplayName.StartsWith(ArgumentStartingValue))
                    {
                        argumentValid = true;
                    }
                }
                if (!argumentValid)
                {
                    lst_Messages.Add(new InspectionMessage()
                    {
                        Message = $"Argument '{argument.DisplayName}' does not start with argument direction"
                    });
                }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    ErrorLevel = configuredRule.ErrorLevel,
                    RecommendationMessage = $"Ensure that argument naming starts with argument direction: {configArgumentStartingValues.Replace("|", ",")}"
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }

        }
        #endregion

        #region 03.1 START OF InspectProperVariableType for ProperVariableTypeRule
        private InspectionResult InspectProperVariableType(IActivityModel activityToInspect, Rule configuredRule)
        {
            var configInvalidTypes = configuredRule.Parameters["Invalid_Types"]?.Value;
            var lst_Messages = new List<InspectionMessage>();

            if (String.IsNullOrWhiteSpace(configInvalidTypes))
            {
                return new InspectionResult() { HasErrors = false };
            }

            foreach (var variable in activityToInspect.Variables)
            {
                foreach (var variableType in configInvalidTypes.Split('|'))
                    if (variable.Type.Split(',')[0].Contains(variableType))
                    {
                        lst_Messages.Add(new InspectionMessage
                        {
                            Message = $"Detected variable/argument with the type of 'Object' or 'GenericValue'. Variable type for: {variable.DisplayName} has type: '{variableType}'.",
                        });
                    }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "It is recommended to specify explicit variable types to enhance workflow clarity and maintainability." +Environment.NewLine+
                                            "Using specific types provides better validation and helps prevent unexpected errors during execution." +Environment.NewLine+
                                            "Please review the variable/argument type and update it to the most appropriate data type for your use case.",
                    ErrorLevel = configuredRule.ErrorLevel
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }
        }
        #endregion

        #region 03.2 START OF InspectProperArgumentType for ProperArgumentTypeRule
        private InspectionResult InspectProperArgumentType(IWorkflowModel workflowToInspect, Rule configuredRule)
        {
            var configInvalidTypes = configuredRule.Parameters["Invalid_Types"]?.Value;
            var lst_Messages = new List<InspectionMessage>();

            if (String.IsNullOrWhiteSpace(configInvalidTypes))
            {
                return new InspectionResult() { HasErrors = false };
            }

            foreach (var argument in workflowToInspect.Arguments)
            {
                foreach (var variableType in configInvalidTypes.Split('|'))
                    if (argument.Type.Split(',')[0].Contains(variableType))
                    {
                        lst_Messages.Add(new InspectionMessage
                        {
                            Message = $"Detected argument with the type of 'Object' or 'GenericValue'. Argument type for: {argument.DisplayName} has type: '{variableType}'.",
                        });
                    }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "It is recommended to specify explicit argument types to enhance workflow clarity and maintainability." +Environment.NewLine+
                                            "Using specific types provides better validation and helps prevent unexpected errors during execution." +Environment.NewLine+
                                            "Please review the argument type and update it to the most appropriate data type for your use case.",
                    ErrorLevel = configuredRule.ErrorLevel
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }
        }
        #endregion

        #region 04.1 START OF InspectActivityRenamed for ActivityRenamedRule
        private InspectionResult InspectActivityRenamed(IActivityModel activityToInspect, Rule configuredRule)
        {
            string activityDisplayName = activityToInspect.DisplayName;
            var lst_Messages = new List<InspectionMessage>();

            string GetDefaultActivityName(string activityTypeAsString)
            {
                try
                {
                    Type type = Type.GetType(activityTypeAsString);
                    DisplayNameAttribute displayNameAttribute = type.IsGenericType ? TypeDescriptor.GetAttributes(type.GetGenericTypeDefinition()).OfType<DisplayNameAttribute>().FirstOrDefault<DisplayNameAttribute>() : TypeDescriptor.GetAttributes(type).OfType<DisplayNameAttribute>().FirstOrDefault<DisplayNameAttribute>();
                    return displayNameAttribute != null ? displayNameAttribute.DisplayName : type.Name.Humanize(LetterCasing.Title);
                }
                catch
                {
                    return (string)null;
                }
            }

            string activityDefaultName = GetDefaultActivityName(activityToInspect.Type);


            if (activityDisplayName == activityDefaultName)
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = $"Detected the use of default display name for an activity: '{activityDisplayName}'."
                });
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "It is important to use meaningful and descriptive names for your activities to enhance workflow readability and maintainability." +Environment.NewLine+
                                            "Using default display names may lead to confusion and make it harder to understand the purpose of each activity.",
                    ErrorLevel = configuredRule.ErrorLevel,
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }

        }

        #endregion

        #region 04.2 START OF InspectAnnotationsInWorkflow for AnnotationsInWorkflowRule
        private InspectionResult InspectAnnotationsInWorkflow(IWorkflowModel workflowToInspect, Rule configuredRule)
        {
            int configMinimumNumberOfAnnotations = int.Parse(configuredRule.Parameters["Minimum_Annotations_Count"]?.Value);
            if (configMinimumNumberOfAnnotations == 0)
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }

            var lst_Messages = new List<InspectionMessage>();

            // Recursive method for getting total number of annotated activities in workflow. 
            int GetNumberOfAnnotatedActivities(IActivityModel rootActivity, int annotationCt)
            {
                

                if (!string.IsNullOrWhiteSpace(rootActivity.AnnotationText))
                {
                    annotationCt++;
                }
                foreach (var activity in rootActivity.Children)
                {
                    
                    annotationCt = GetNumberOfAnnotatedActivities(activity, annotationCt);
                }
                return annotationCt;
            }

            // Get total number of annotated activities in workflow
            int AnnotationsInWorkflow = GetNumberOfAnnotatedActivities(workflowToInspect.Root, 0);

            // Evaluate result
            if (AnnotationsInWorkflow >= configMinimumNumberOfAnnotations)
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }
            else
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = $"The workflow '{workflowToInspect.DisplayName}' contains {AnnotationsInWorkflow.ToString()} annotations, which is below the configured minimum count of {configMinimumNumberOfAnnotations.ToString()}."
                });
                return new InspectionResult()
                {
                    HasErrors = true,
                    ErrorLevel = configuredRule.ErrorLevel,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "Annotations are essential for documenting your workflow, providing context, and making it more understandable for both you and your team." +Environment.NewLine+
                                            "Consider adding additional annotations to explain key steps, decisions, or any other important information."
                };
            }
        }
        #endregion

        #region 04.3 START OF InspectLogsInWorkflow for LogMsgInWorkflowRule
        private InspectionResult InspectLogsInWorkflow(IWorkflowModel workflowToInspect, Rule configuredRule)
        {
            int configMinimumNumberOfLogs = int.Parse(configuredRule.Parameters["Minimum_Logs_Count"]?.Value);
            var lst_Messages = new List<InspectionMessage>();
            // Recursive method for getting total number of log message activities in workflow. 
            int GetNumberOfLogs(IActivityModel rootActivity, int LogsCt)
            {

                if (rootActivity.ToolboxName == "LogMessage")
                {
                    LogsCt++;
                }
                foreach (var activity in rootActivity.Children)
                {
                    LogsCt = GetNumberOfLogs(activity, LogsCt);
                }
                return LogsCt;
            }

            // Get number of log message activites in workflow
            int LogsCount = GetNumberOfLogs(workflowToInspect.Root, 0);

            // Evaluate result
            if (LogsCount >= configMinimumNumberOfLogs)
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }
            else
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = $"The workflow '{workflowToInspect.DisplayName}' contains {LogsCount.ToString()} log messages, which is below the configured minimum count of {configMinimumNumberOfLogs.ToString()}."
                });
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    ErrorLevel = configuredRule.ErrorLevel,
                    RecommendationMessage = "Logging is crucial for capturing important information during workflow execution." +Environment.NewLine+
                                            "Consider adding additional log messages to record key events, decisions, or any other significant details."
                };
            }
        }
        #endregion

        #region 05 START OF InspectFrameworkType for UsingREFRule
        private InspectionResult InspectFrameworkType(IProjectModel projectToInspect, Rule configuredRule)
        {
            List<string> lst_REFWorkflows = new List<string>()
            {
                "InitAllSettings",
                "InitAllApplications",
                "KillAllProcesses",
                "CloseAllApplications"
            };
            var lst_Messages = new List<InspectionMessage>();
            var lst_WorkflowsInProject = new List<string>();
            bool IsREF = true;


            foreach (var workflow in projectToInspect.Workflows)
            {
                lst_WorkflowsInProject.Add(workflow.DisplayName);
            }

            foreach (var workflow in lst_REFWorkflows)
            {
                if (!lst_WorkflowsInProject.Contains(workflow))
                {
                    IsREF = false;
                    break;
                }
            }
            if (!IsREF)
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = "The workflow does not appear to be utilizing the REFramework."
                });
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    ErrorLevel = configuredRule.ErrorLevel,
                    RecommendationMessage = "The Robotic Enterprise Framework (REFramework) is a powerful template in UiPath designed for building scalable and maintainable automation solutions." +Environment.NewLine+
                                            "It enforces best practices, promotes code reusability, and enhances error handling."
                };
            }
            else
            {
                return new InspectionResult() { HasErrors = false };
            }


        }
        #endregion

        #region 06.1 START OF InspectPasswordVariables for PasswordSecureStringRule
        private InspectionResult InspectPasswordVariables(IActivityModel activityToInspect, Rule configuredRule)
        {
            var configPasswordVarContains = configuredRule.Parameters["Password_Variable_Contains"]?.Value;
            var lst_Messages = new List<InspectionMessage>();

            foreach (var variable in activityToInspect.Variables)
            {
                var varIsPass = false;
                foreach (var PasswordVarContains in configPasswordVarContains.Split('|'))
                {
                    if (variable.DisplayName.ToUpper().Contains(PasswordVarContains.ToUpper()))
                    {
                        varIsPass = true;
                        break;
                    }
                }

                if (varIsPass)
                {
                    Type varType = Type.GetType(variable.Type);

                    if (!varType.Name.Equals("SecureString"))
                    {
                        lst_Messages.Add(new InspectionMessage()
                        {
                            Message = $"The workflow contains a password variable '{variable.DisplayName}' that is not of type 'SecureString'."
                        });
                    }
                }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    ErrorLevel = configuredRule.ErrorLevel,
                    InspectionMessages = lst_Messages,
                    HasErrors = true,
                    RecommendationMessage = "For security reasons, it is mandatory to use the 'SecureString' data type for storing password variables."+Environment.NewLine+
                                            "Using 'SecureString' helps protect sensitive information by encrypting it in memory and providing additional security against potential vulnerabilities."
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false
                };
            }
        }
        #endregion

        #region 06.2 START OF InspectPasswordInputDialogVariables for InputDialogIsPasswordRule
        private InspectionResult InspectPasswordInputDialogVariables(IActivityModel activityToInspect, Rule configuredRule)
        {
            var lst_Messages = new List<InspectionMessage>();
            Type activityType = Type.GetType(activityToInspect.Type);
            string passwordArgName = null;
            bool isSecureString = false;
            bool isPasswordProperty = false;


            if (activityType.Name.Equals("InputDialog"))
            {
                foreach (var argument in activityToInspect.Arguments)
                {
                    if (argument.DisplayName.Equals("Result"))
                    {
                        passwordArgName = argument.DefinedExpression;
                    }
                }

                foreach (var property in activityToInspect.Properties)
                {
                    if (property.DisplayName.Equals("IsPassword") && bool.Parse(property.DefinedExpression))
                    {
                        isPasswordProperty = true;
                    }
                }

                foreach (var variable in activityToInspect.Parent.Variables)
                {
                    if (variable.DisplayName.Equals(passwordArgName))
                    {
                        Type varType = Type.GetType(variable.Type);


                        if (varType.Name.Equals("SecureString"))
                        {
                            isSecureString = true;
                        }
                    }
                }

                if (isSecureString && !isPasswordProperty)
                {
                    lst_Messages.Add(new InspectionMessage()
                    {
                        Message = $"Input dialog {activityToInspect.DisplayName} Result argument '{passwordArgName}' is a 'SecureString' variable type, but 'IsPassword' property is not set to 'true'."
                    });
                }
                else if (isPasswordProperty && !isSecureString)
                {
                    lst_Messages.Add(new InspectionMessage()
                    {
                        Message = $"Input dialog {activityToInspect.DisplayName} 'IsPassword' property is set to 'true', but Result argument '{passwordArgName}' is not a 'SecureString' variable type."
                    });
                }
            }


            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "Ensuring secure handling of passwords is crucial for protecting sensitive information." +Environment.NewLine+
                                            "When using an Input Dialog for password input, follow best practices by setting the 'IsPassword' property to true" +Environment.NewLine+
                                            "and ensuring that the output result variable is of type 'SecureString'.",
                    ErrorLevel = configuredRule.ErrorLevel,
                };
            }
            else
            {
                return new InspectionResult() { HasErrors = false };
            }
        }
        #endregion

        #region 07 START OF InspectNoDelayActivity for NoDelayActivityRule
        private InspectionResult InspectNoDelayActivity(IActivityModel activityToInspect, Rule configuredRule)
        {
            var lst_Messages = new List<InspectionMessage>();
            Type activityType = Type.GetType(activityToInspect.Type);


            if (activityType.Name.Equals("Delay"))
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = $"The workflow contains the use of Delay activities, which is discouraged."
                });
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "Using Delay activities can introduce unnecessary delays and make your workflows less efficient." +Environment.NewLine+
                                            "Instead, leverage activities that are specifically designed for waiting, such as Check App State or Element Exists.",
                    ErrorLevel = configuredRule.ErrorLevel,
                };
            }
            else
            {
                return new InspectionResult() { HasErrors = false };
            }
        }
        #endregion

        #region 08 START OF InspectExcelVisible for ExcelVisibleRule
        private InspectionResult InspectExcelVisible(IActivityModel activityToInspect, Rule configuredRule)
        {
            var lst_Messages = new List<InspectionMessage>();
            foreach (var property in activityToInspect.Properties)
            {
                // For Modern Excel Process Scopes
                if (property.DisplayName.Equals("Show Excel window", StringComparison.CurrentCultureIgnoreCase))
                {
                    try
                    {
                        if (bool.Parse(property.DefinedExpression)) // if property Show Excel Window is set to True
                        {
                            lst_Messages.Add(new InspectionMessage()
                            {
                                Message = $"The workflow uses Excel activities without the Excel Application Scope in the background. Excel Process Scope '{activityToInspect.DisplayName}' has property '{property.DisplayName}' set to 'True'",
                            });
                        }
                    }
                    catch // if set value of property cannot be read, means user did not specify property value
                    {
                        lst_Messages.Add(new InspectionMessage()
                        {
                            Message = $"The workflow uses Excel activities without the Excel Application Scope in the background. Excel Process Scope '{activityToInspect.DisplayName}' has property '{property.DisplayName}' set to 'Same as project'"
                        });
                    }
                }

                //For Classic Excel Application Scopes
                if (property.DisplayName.Equals("Visible", StringComparison.CurrentCultureIgnoreCase) && bool.Parse(property.DefinedExpression))
                {
                    lst_Messages.Add(new InspectionMessage()
                    {
                        Message = $"The workflow uses Excel activities without the Excel Application Scope in the background." +Environment.NewLine+
                                   "xcel Application Scope '{activityToInspect.DisplayName}' has property '{property.DisplayName}' set to 'True'"
                    });
                }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    ErrorLevel = configuredRule.ErrorLevel,
                    RecommendationMessage = "It is recommended to use the Excel Application Scope activity when working with Excel files to ensure better performance and reliability." +Environment.NewLine+
                                            "Additionally, set the 'Show Excel Window'/'Visible' property to 'False' within the Excel Application Scope to run Excel operations in the background without displaying the Excel application window."
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false
                };
            }
        }
        #endregion

        #region 09 START OF InspectOutlookGetMail for OutlookFifoRule
        private InspectionResult InspectOutlookGetMail(IActivityModel activityToInspect, Rule configuredRule)
        {
            var lst_Messages = new List<InspectionMessage>();

            if (activityToInspect.ToolboxName.Equals("GetOutlookMailMessages", StringComparison.InvariantCultureIgnoreCase))
            {
                foreach (var property in activityToInspect.Properties)
                {
                    if (property.DisplayName.Equals("OrderByDate", StringComparison.InvariantCultureIgnoreCase) && !property.DefinedExpression.Equals("OldestFirst", StringComparison.InvariantCultureIgnoreCase))
                    {
                        lst_Messages.Add(new InspectionMessage()
                        {
                            Message = $"The workflow uses the Get Outlook Mail Messages activity: '{activityToInspect.DisplayName}' without setting the OrderByDate property to 'OldestFirst', potentially leading to non-FIFO mail processing."
                        });
                    }
                }
            }


            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    ErrorLevel = configuredRule.ErrorLevel,
                    RecommendationMessage = "To ensure that Outlook mails are processed in a First-In-First-Out (FIFO) manner," +Environment.NewLine+
                                            "it is recommended to set the OrderByDate property of the Get Outlook Mail Messages activity to 'OldestFirst'." +Environment.NewLine+
                                            "This ensures that the oldest emails are processed first."
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false
                };
            }
        }
        #endregion

        #region 10 START OF InspectOutlookGetMailFilters for OutlookFiltersRule
        private InspectionResult InspectOutlookGetMailFilters(IActivityModel activityToInspect, Rule configuredRule)
        {
            var lst_Messages = new List<InspectionMessage>();
            Type activityType = Type.GetType(activityToInspect.Type);
            bool filterUsed = false;
            bool filterByIdUsed = false;


            if (activityType.Name.Equals("GetOutlookMailMessages"))
            {
                foreach (var argument in activityToInspect.Arguments)
                {
                    if (argument.DisplayName.Equals("Filter") && !string.IsNullOrEmpty(argument.DefinedExpression))
                    {
                        filterUsed = true;
                    }
                    else if (argument.DisplayName.Equals("FilterByMessageIds") && !string.IsNullOrEmpty(argument.DefinedExpression))
                    {
                        filterByIdUsed = true;
                    }
                }
            }
            else 
            {
                return new InspectionResult()
                {
                    HasErrors = false
                };
            }

            if (!(filterUsed || filterByIdUsed))
            {
                lst_Messages.Add(new InspectionMessage()
                {
                    Message = "The workflow uses the Get Outlook Mail Messages activity without applying a filter to query mail messages."
                });
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    ErrorLevel = configuredRule.ErrorLevel,
                    RecommendationMessage = "It is recommended to apply a filter when querying mail messages using the Get Outlook Mail Messages activity." +Environment.NewLine+
                                            "This allows you to narrow down the search and retrieve only the relevant emails," +Environment.NewLine+
                                            "improving performance and ensuring that your automation processes the required data."
                };
            }
            else
            {
                return new InspectionResult()
                {
                    HasErrors = false,
                };
            }
            

        }
        #endregion

        #region 11 START OF InspectSensitiveNotLogged for SensitiveNotLogRule
        private InspectionResult InspectSensitiveNotLogged(IActivityModel activityToInspect, Rule configuredRule)
        {
            String configureVariableNames = configuredRule.Parameters["Variable_Names"]?.Value;
            String[] variableNames = configureVariableNames.Split('|');
            if (String.IsNullOrEmpty(configureVariableNames))
            {
                return new InspectionResult() { HasErrors = false };
            }

            var lst_Messages = new List<InspectionMessage>();

            String RegexValue = @"""([^""]*)""";

            if (activityToInspect.ToolboxName.Equals("LogMessage") || activityToInspect.ToolboxName.Equals("WriteLine"))
            {
                
                foreach (var argument in activityToInspect.Arguments)
                {
                    var value = argument.DefinedExpression;

                    if (String.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    String parsedValue = Regex.Replace(value.ToString().ToLower(), RegexValue, String.Empty);

                    String[] arrayValue = parsedValue.Split('|');

                    for (var i = 0; i < variableNames.Length; i++)
                    {
                        for (var j = 0; j < arrayValue.Length; j++)
                        {

                            if (arrayValue[j].Contains(variableNames[i].ToLower()))
                            {
                                lst_Messages.Add(new InspectionMessage()
                                {
                                    Message = $"The workflow contains Log Message or Write Line activity '{activityToInspect.DisplayName}' that may output sensitive information from {variableNames[i]}."                                });
                            }
                        }
                    }
                }
            }

            if (lst_Messages.Count > 0)
            {
                return new InspectionResult()
                {
                    HasErrors = true,
                    InspectionMessages = lst_Messages,
                    RecommendationMessage = "It is crucial to avoid logging or displaying sensitive information using Log Message or Write Line activities," +Environment.NewLine+
                                            "as this can pose a security risk." +Environment.NewLine+
                                            "Ensure that your automation processes sensitive data securely and follows best practices for handling confidential information.",
                    ErrorLevel = configuredRule.ErrorLevel
                };
            }
            else
            {
                return new InspectionResult() { HasErrors = false };
            }
        }
        #endregion
    }
}
