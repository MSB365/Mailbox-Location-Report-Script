# Exchange Online Mailbox Location Report Generator

A PowerShell script that generates comprehensive HTML reports showing all Exchange mailboxes and their locations (Exchange Online vs On-Premises) in hybrid environments.

## Features

- **Complete Mailbox Inventory**: Displays all mailbox types (User, Shared, Room, Equipment, etc.)
- **Location Detection**: Accurately identifies whether mailboxes are in Exchange Online or On-Premises
- **Interactive HTML Report**: Filterable report with search and sorting capabilities
- **Summary Statistics**: Shows counts and percentages by location and type
- **Color-Coded Display**: Visual indicators for easy identification
- **Hybrid Environment Support**: Works with both pure cloud and hybrid Exchange deployments

## Prerequisites

- PowerShell 5.1 or later
- Exchange Online Management module
- Appropriate Exchange Online permissions (View-Only Organization Management or higher)

## Installation

1. Install the Exchange Online Management module:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force
