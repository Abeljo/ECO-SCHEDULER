# Pest Control Scheduler

Automated scheduling system for March pest control visits with intelligent load balancing and Excel export.

## Features
- **Deterministic Scheduling**: Strictly follows frequency rules (Monthly, Bi-Monthly, Quarterly, VIP).
- **Load Balancing**: Distributes jobs evenly across 4 technician teams to ensure daily continuity.
- **Excel Export**: Generates a formatted, color-coded `March_Pest_Control_Schedule.xlsx`.
- **Validation**: Automatically checks for rule violations (spacing, capacity).

## Setup & Usage
1. **Install Dependencies**:
   ```bash
   npm install
   ```
2. **Input Data**: Ensure `customers.json` is in the root directory.
3. **Run Scheduler**:
   ```bash
   node generateSchedule.js
   ```

## Scheduling Rules
- **Monthly**: 2 visits (12–16 days apart).
- **Bi-Monthly**: 4 visits (~7–8 days apart).
- **VIP (EDR)**: Every 2 days (including Sundays).
- **Capacity**: Target max 18 jobs per day.
- **Work Days**: Monday–Saturday (Sundays reserved for VIP).
