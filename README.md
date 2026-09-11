# WhatsApp-AI-Agent
WhatsApp AI Assistant

A full-featured WhatsApp AI Assistant that connects WhatsApp with Microsoft services and custom Python modules to provide employee support, company knowledge access, workflow automation, reminders, calendar management, request tracking, task management, expense tracking, speech support, and long-term conversational memory.

This project was developed progressively, beginning with a basic WhatsApp webhook and message handling system and expanding into a modular AI assistant capable of interacting with SharePoint, Microsoft Graph, Copilot Studio-related workflows, Azure services, and internal business data.

Project Overview

The goal of this project is to provide users with a convenient way to access company services and information directly from WhatsApp.

Instead of opening multiple systems or applications, a user can send a natural-language message to the assistant and receive help with tasks such as:

Asking questions about approved company documents

Submitting a company request

Checking the status of a previous request

Creating and managing tasks

Viewing upcoming tasks

Creating calendar events

Checking for calendar conflicts

Viewing today's or tomorrow's calendar

Receiving automated reminders

Snoozing or dismissing reminders

Recording expenses

Viewing monthly spending

Comparing spending against budgets

Accessing personalized employee information

Using speech and audio features

Maintaining short-term and long-term conversational memory

Handling errors and user feedback cleanly

Showing typing indicators and response timing

Processing multiple types of requests through one WhatsApp interface

Development Journey

The project evolved through several stages.

1. WhatsApp Webhook Integration

The first stage focused on connecting WhatsApp to a Python backend.

A Flask webhook was created to:

Verify the Meta webhook

Receive incoming WhatsApp messages

Extract the sender's WhatsApp number

Read message content

Route messages to the assistant

Send replies back through the WhatsApp Cloud API

This provided the foundation for the entire project.

2. Microsoft and SharePoint Integration

The next stage connected the assistant to Microsoft services.

The assistant was expanded to work with:

Microsoft SharePoint

Microsoft Graph API

Microsoft Entra ID

Company documents

SharePoint lists

Employee data

Internal business information

The assistant could retrieve information from approved company sources rather than relying only on general AI knowledge.

3. Permission-Aware Employee Access

A major feature of the project was permission-aware access.

The assistant was designed so that sensitive information is returned only to the correct user.

For example, the assistant can retrieve employee-specific information such as:

Name

Department

Job title

Salary

Bonus

Currency

Other approved employee information

The user's identity is matched to organizational data so that one employee cannot retrieve another employee's private information.

4. Company Document Question Answering

The assistant was expanded to answer questions using approved company documents.

This allows users to ask natural-language questions such as:

What is the company's remote work policy?

or:

How many vacation days are employees allowed?

The assistant searches organizational content and provides responses based on approved information.

This creates a Retrieval-Augmented Generation style workflow where responses are grounded in company data.

5. Employee Request Management

A request management feature was added using SharePoint.

Users can submit requests through WhatsApp and later retrieve their status.

Example:

add request

The assistant can create a new request in a SharePoint list.

Users can later ask:

check my request

The assistant retrieves the request that belongs to the current user and returns its status.

This functionality is handled through the request client module.

6. Task Management

The assistant was expanded into a personal task manager.

Users can create and manage tasks directly from WhatsApp.

Supported task capabilities include:

Create tasks

Add due dates

Set priorities

Retrieve all tasks

Retrieve a specific user's task

Retrieve tasks due on a specific date

Mark tasks as complete

Delete tasks

Associate tasks with the user's WhatsApp number

Return task links and task context

Tasks are stored in the SharePoint list:

WhatsApp Tasks

Each task can contain information such as:

Title
DueDate
Priority
Status
WhatsAppNumber

7. Calendar Integration

Microsoft Graph calendar integration was added to allow users to manage their schedules through WhatsApp.

Capabilities include:

Retrieve today's calendar events

Retrieve tomorrow's events

Retrieve events for a specific date

Create calendar events

Extract event details from natural language

Check for scheduling conflicts

Compare event start and end times

Suggest available scheduling times

The system uses the timezone:

Asia/Beirut

Conflict detection compares proposed event times with existing Microsoft Outlook calendar events.

8. Calendar Conflict Detection

Before creating or rescheduling an event, the assistant can determine whether the requested time overlaps with an existing calendar event.

The overlap logic checks whether two time ranges intersect.

This allows the assistant to prevent double booking and provide more useful scheduling assistance.

9. Automated Calendar Reminders

A reminder scheduler was added to monitor upcoming events.

The reminder system:

Checks for events starting soon

Sends a WhatsApp reminder before the event

Tracks reminder state

Prevents duplicate reminders

Allows the user to snooze a reminder

Allows the user to dismiss a reminder

Current reminder behavior includes:

Reminder lead time: 15 minutes
Snooze duration: 5 minutes
Scheduler check interval: 60 seconds

This transformed the assistant from a reactive chatbot into a system that can proactively notify the user.

10. Expense Tracking

Expense management was added so users can record personal or business expenses through WhatsApp.

Users can enter information such as:

Expense title

Amount

Category

Currency

Expense date

The assistant associates expenses with the user's WhatsApp number.

Example data structure:

Title
Amount
Category
Currency
ExpenseDate
WhatsAppNumber

11. Monthly Expense Summaries

The assistant can calculate monthly expense totals.

It can also organize totals by category.

For example:

Food: $250
Transportation: $140
Office: $320

This provides users with a simple financial summary directly inside WhatsApp.

12. Budget Tracking

Budget functionality was added on top of expense tracking.

The assistant can retrieve category budgets and compare them with spending.

Example:

Category: Transportation
Budget: $300
Spent: $210
Remaining: $90

This allows the system to function as a simple conversational budget assistant.

13. Short-Term Conversation History

The assistant stores recent messages so that it can maintain conversational context.

This makes follow-up questions possible without requiring the user to repeat all previous information.

Recent history is maintained separately from long-term semantic memory.

14. Vector-Based Long-Term Memory

A vector memory system was added to allow the assistant to recall older but semantically relevant information.

The memory system includes two layers:

Recent Memory

Stores the most recent messages for immediate context.

Semantic Memory

Stores embedded versions of earlier messages and retrieves relevant information using cosine similarity.

Important configuration used in the project includes:

TOP_K = 3
SIMILARITY_THRESHOLD = 0.30
MIN_TEXT_LENGTH = 15
RECENT_WINDOW = 6

The system can:

Store user messages

Store assistant messages

Create embeddings

Save memory to persistent storage

Calculate cosine similarity

Retrieve relevant previous information

Build memory context for future responses

Avoid duplicating very recent messages in semantic recall

Memory is stored in:

vector_memory.json

15. Multiple Request Handling

The assistant was designed to support messages containing more than one request.

For example:

I need a replacement monitor because my current one is damaged, and I would also like to work remotely from September 20 to September 22.

The system can identify that this contains multiple intents and route the information to the appropriate workflow.

This became an important part of the project's later development.

16. Workflow Routing

The assistant uses dedicated modules to separate responsibilities.

Instead of placing all business logic in one file, functionality was divided into specialized clients and helper modules.

This improves:

Code organization

Maintainability

Debugging

Testing

Reuse

Separation of concerns

17. Typing Indicator

A typing indicator feature was added to improve user experience.

The assistant can signal that it is processing a request before sending the final response.

This makes the WhatsApp interaction feel more natural.

18. Response Timing

A response timer module was added to measure and manage response timing.

This can be used to:

Track processing duration

Monitor assistant responsiveness

Help identify slow operations

Improve the overall user experience

19. Error Handling

A dedicated error handling module was added to manage exceptions and unexpected failures.

Instead of exposing technical errors directly to the user, the assistant can return a clean and understandable response.

This improves reliability and prevents the application from failing silently.

20. Speech Support

Speech functionality was added through a dedicated speech client.

This allows the architecture to support voice-related interaction and audio processing as part of the WhatsApp experience.

Current Capabilities

The WhatsApp AI Assistant is currently capable of acting as a multi-purpose workplace assistant.

It combines conversational AI with business systems to provide the following capabilities:

Company Knowledge

Search approved company documents

Answer policy and organizational questions

Ground responses in internal information

Employee Services

Retrieve user-specific information

Apply permission-aware access

Protect private employee information

Requests

Submit requests

Retrieve request status

Associate requests with individual users

Tasks

Create tasks

View tasks

Retrieve tasks due on a date

Complete tasks

Delete tasks

Manage task priority and status

Calendar

View events

Create events

Detect schedule conflicts

Work with specific dates

Handle scheduling requests

Reminders

Detect upcoming events

Send automatic WhatsApp notifications

Snooze reminders

Dismiss reminders

Expenses

Record expenses

Categorize spending

Calculate monthly totals

Display spending by category

Budgets

Retrieve category budgets

Compare budgets with expenses

Calculate remaining budget

AI and Conversation

Natural-language interaction

Short-term memory

Vector-based long-term memory

Semantic recall

Multi-intent handling

Context-aware responses

User Experience

WhatsApp-native interaction

Typing indicators

Response timing

Error handling

Speech support

Project Files

The project is organized into multiple Python modules, each responsible for a specific part of the assistant.

.
├── app.py
├── azure_client.py
├── commands.py
├── error_handler.py
├── response_timer.py
├── typing_indicator.py
├── documents.py
├── sharepoint.py
├── calendar_client.py
├── expense_client.py
├── reminders.py
├── request_client.py
├── task_client.py
├── speech_client.py
├── whatsapp.py
├── vector_memory.py
└── __pycache__/

File Responsibilities

app.py

Main application entry point.

Responsible for:

Starting the Flask application

Receiving webhook requests

Coordinating message processing

Connecting the major components of the assistant

whatsapp.py

Handles WhatsApp-specific communication.

Responsibilities include:

Processing incoming WhatsApp messages

Extracting sender information

Sending responses

Communicating with the Meta WhatsApp Cloud API

azure_client.py

Handles Azure-related AI or service integration used by the assistant.

It provides the connection between the application and Azure-backed functionality.

commands.py

Contains command-processing and routing logic.

It helps determine which feature should handle a user's request.

documents.py

Handles company document processing and information retrieval.

Used for answering questions from approved organizational content.

sharepoint.py

Provides SharePoint connectivity and shared SharePoint operations.

Used by the assistant to interact with company data and SharePoint resources.

request_client.py

Handles employee request workflows.

Capabilities include:

Creating requests

Retrieving user requests

Checking request status

task_client.py

Handles task management.

Capabilities include:

Creating tasks

Retrieving tasks

Retrieving user-specific tasks

Retrieving tasks by date

Completing tasks

Deleting tasks

Returning task context and links

calendar_client.py

Handles Microsoft Graph calendar functionality.

Capabilities include:

Event retrieval

Event detail extraction

Conflict checking

Time overlap detection

Date-based calendar lookup

reminders.py

Handles automated calendar reminders.

Capabilities include:

Finding events starting soon

Scheduling reminder checks

Sending reminders

Snoozing reminders

Dismissing reminders

Tracking reminder state

expense_client.py

Handles expenses and budget information.

Capabilities include:

Adding expenses

Calculating monthly totals

Grouping expenses by category

Retrieving category budgets

vector_memory.py

Provides long-term semantic memory.

Capabilities include:

Storing conversations

Creating embeddings

Saving memory

Loading memory

Cosine similarity

Semantic recall

Recent-message exclusion

Context generation

speech_client.py

Handles speech and audio-related assistant capabilities.

typing_indicator.py

Controls typing indicators to improve the WhatsApp user experience.

response_timer.py

Tracks assistant response time and processing duration.

error_handler.py

Provides centralized error handling and user-friendly failure responses.

High-Level Architecture

                      ┌──────────────────────┐
                      │       User           │
                      └──────────┬───────────┘
                                 │
                                 ▼
                      ┌──────────────────────┐
                      │      WhatsApp        │
                      └──────────┬───────────┘
                                 │
                                 ▼
                ┌──────────────────────────────┐
                │ Meta WhatsApp Cloud API      │
                └──────────────┬───────────────┘
                               │
                               ▼
                    ┌────────────────────┐
                    │ Flask Webhook      │
                    │      app.py        │
                    └─────────┬──────────┘
                              │
                              ▼
                    ┌────────────────────┐
                    │ Command Routing    │
                    │    commands.py     │
                    └─────────┬──────────┘
                              │
       ┌──────────────────────┼────────────────────────┐
       │                      │                        │
       ▼                      ▼                        ▼
┌──────────────┐      ┌──────────────┐        ┌──────────────┐
│ SharePoint   │      │ Graph API    │        │ Azure / AI   │
│ Documents    │      │ Calendar     │        │ Services     │
│ Requests     │      │ Events       │        │ Memory       │
│ Tasks        │      │ Reminders    │        │ Speech       │
│ Expenses     │      └──────────────┘        └──────────────┘
└──────────────┘

Technology Stack

The project uses:

Python

Flask

Microsoft SharePoint

Microsoft Graph API

Microsoft Entra ID

Azure services

Microsoft Copilot Studio-related workflows

Meta WhatsApp Cloud API

REST APIs

ngrok

JSON

Vector embeddings

Cosine similarity

Semantic search

SharePoint lists

Microsoft Outlook Calendar

Example User Interactions

Company Information

User:
What is the company's remote work policy?

Assistant:
Returns information based on approved company documents.

Request Management

User:
Add a request for a replacement monitor.

Assistant:
Creates the request and confirms submission.

User:
Check my request.

Assistant:
Retrieves the request associated with the user's WhatsApp number.

Tasks

User:
Remind me to submit my report tomorrow.

Assistant:
Creates a task with the appropriate due date.

Calendar

User:
Schedule my assignment presentation today from 3:30 PM to 4:30 PM.

Assistant:
Checks the calendar for conflicts before creating the event.

Multiple Requests

User:
I need a replacement monitor and I want to work remotely from September 20 to September 22.

Assistant:
Identifies the separate requests and routes them to the appropriate workflows.

Expenses

User:
Add a $45 transportation expense for today.

Assistant:
Stores the expense under the user's account.

Budget

User:
How much of my transportation budget is left?

Assistant:
Compares the user's spending with the configured transportation budget.

Outcomes

This project successfully demonstrates how WhatsApp can be transformed into a practical enterprise AI interface.

Major outcomes include:

Successfully connected WhatsApp to a Python Flask backend

Successfully verified and operated a Meta WhatsApp webhook

Successfully sent and received WhatsApp messages through the Cloud API

Connected the assistant to Microsoft SharePoint

Retrieved approved organizational information

Implemented permission-aware access to employee data

Added SharePoint-based request tracking

Added task management

Added Microsoft Graph calendar integration

Added event conflict detection

Added automated event reminders

Added snooze and dismiss reminder actions

Added expense tracking

Added budget tracking

Added short-term conversation history

Added vector-based long-term memory

Added semantic recall

Added multiple-request handling

Added modular command routing

Added typing indicators

Added response-time handling

Added centralized error handling

Added speech-related functionality

Created a modular architecture that can be extended with additional enterprise workflows

Security and Privacy

Because the assistant interacts with organizational information, security is an important part of the design.

The project follows these principles:

Sensitive employee information should only be returned to the correct user

Organizational data should come from approved sources

Credentials should never be hard-coded

API keys and secrets should be stored in environment variables

SharePoint permissions should be respected

Microsoft Graph permissions should follow least-privilege principles

WhatsApp access tokens should remain private

Private .env files should never be committed to GitHub

Employee information should not be exposed in logs

Environment Variables

A typical .env configuration may contain values such as:

WHATSAPP_ACCESS_TOKEN=
WHATSAPP_PHONE_NUMBER_ID=
WHATSAPP_VERIFY_TOKEN=

MICROSOFT_TENANT_ID=
MICROSOFT_CLIENT_ID=
MICROSOFT_CLIENT_SECRET=

SHAREPOINT_SITE_ID=
SHAREPOINT_LIST_ID=

AZURE_ENDPOINT=
AZURE_API_KEY=

Never commit real secrets to GitHub.

Installation

Clone the repository:

git clone <your-repository-url>
cd <repository-folder>

Create a virtual environment:

python -m venv venv

Activate it.

Windows

venv\Scripts\activate

macOS / Linux

source venv/bin/activate

Install dependencies:

pip install -r requirements.txt

Running the Application

Start the Flask application:

python app.py

For local webhook testing, start ngrok:

ngrok http 5000

Use the generated HTTPS URL as the callback URL in the Meta Developer configuration.

Example:

https://your-ngrok-url.ngrok-free.app/webhook

Suggested .gitignore

The repository should exclude temporary files, credentials, cache files, and local memory data.

__pycache__/
*.pyc

.env
venv/
.venv/

vector_memory.json

.DS_Store

Future Development

The architecture supports continued expansion.

Possible future improvements include:

More advanced multi-intent processing

More autonomous workflow execution

Additional Microsoft 365 integrations

Teams integration

Outlook email integration

Advanced document retrieval

Improved memory ranking

Memory management controls

Admin dashboard

Analytics and usage monitoring

More proactive notifications

More sophisticated scheduling suggestions

Expanded speech and voice support

Improved authentication

More detailed permission rules

Additional SharePoint business workflows

Project Purpose

The purpose of this project is to demonstrate how conversational AI, WhatsApp, Microsoft 365, SharePoint, Microsoft Graph, Azure, automation, and long-term memory can be combined into one practical assistant.

The result is more than a chatbot. It is a modular business assistant capable of interacting with users, organizational data, workflows, schedules, tasks, requests, expenses, reminders, and memory through a single WhatsApp conversation.

Author

Jovita Issa
