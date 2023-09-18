# Coding Challenge
This project includes the necessary package to implement the scenario provided as part of the coding challenge.

# How to build the code
1. Clone the repository to your local machine.
2. Update or manage NuGet packages if you encounter reference issues during the build process for D365.Workflow.TFL (code activity).

# How to Run or Test
Please refer to the document in the 'Document' folder for instructions on solution import and Test-Driven Development (TDD). Additionally, use the solution in the 'Managed Solution' folder to import into PowerApps.

# Assumption
I have built the solution assuming that each business unit has its own manager, escalation team, and confidential team. All users of the business unit, including agents, should have visibility into cases and follow-up activities for better tracking of cases created against contacts

To achieve this, I've created and utilized an Access team to share confidential email activities with the Confidential Case Team. As part of this process, I've established a 'Parent Confidential' team at the parent Business Unit (BU) level. The custom code development will specifically share emails with members of the Confidential Case Team (e.g., Confidential Case Team - Underground) and manager. Additionally, there are two different email templates: one for the confidential team and another for the escalation team. These templates contain dynamic values for contact and case details, and they are configurable and can be edited to accommodate business needs.
