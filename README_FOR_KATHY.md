# CMOA Tour Scheduler — Quick Guide for Kathy

Hi Kathy! This is a short guide to the new tour scheduling system. The good news:
**most things now happen automatically.** Docents sign up on the website, the
system assigns them fairly, sends them reminder emails and calendar invites,
and finds substitutes when someone drops out. You mostly just watch it work.

---

## The Website

**https://svijayakumar2.github.io/cmoa-docent-scheduling/**

You don't need to do anything on the website, but it's the easiest way to
*see* the schedule:

1. Open the link above.
2. Click **"Just want to see the schedule? View the monthly calendar"**
   (no need to type a name).
3. You'll see a calendar that looks like a wall calendar — every tour, with
   the assigned docents' names on it. Use **Prev / Next** to change months.

### Printing the calendar for the docent room

On the monthly calendar, click the **"Print This Calendar"** button.
It prints sideways (landscape) so the whole month fits on one page —
ready to post on the board. Print a fresh copy whenever you like;
it always shows the latest assignments.

---

## The Spreadsheet

The spreadsheet is the "brain" of the system. It has 4 tabs at the bottom.
Here is what is safe to change and what is not.

### Tab 1: "Schedule" (the list of tours)

**Things you SHOULD do here:**

| What you want to do | How |
|---|---|
| Add a new tour | Add a new row and fill in columns **A–E** (Slot ID, Date, Time, Tour Type, # Docents Needed). Column **H** (Details) is optional. |
| Cancel a tour | In column **F** (Status), type exactly: **Tour Cancelled** |
| A docent told you they can't make it | In column **F** (Status), type exactly: **Needs Sub** |
| The system emails you "please assign manually" | Type the docent's name in column **G** (Assigned) and set column **F** to **Assigned** |

**Important:** the moment you type "Tour Cancelled" or "Needs Sub" in
column F, the system **immediately emails docents**. Only type it when
you mean it.

**Things you should NOT do here:**

- Do **not** delete rows — to cancel a tour, use "Tour Cancelled" instead.
- Do **not** type in column **G** (Assigned) except when the system asks
  you to assign manually. The system fills this in on its own.
- Do **not** change column **A** (Slot ID) on existing rows.

### Tab 2: "Docents" (the list of docents)

**Things you SHOULD do here:**

- Add a new docent: fill in their **Name** (column A), **Email** (column B),
  and **Certified Tours** (column F — for example: `pc, ci, sch`).
- Mark whether they can lead school tours: column **I** (Lead Eligible) —
  type **Yes** or leave blank.
- Fix a wrong email address.

**Things you should NOT touch here:**

- Columns **C** and **D** (tour counts) — the system counts these itself.
- Columns **E, G, H, J** (vacations and day preferences) — docents set
  these themselves on the website. You *can* fill them in if a docent
  calls you and asks, but normally leave them alone.

### Tab 3: "Signups" — DO NOT TOUCH

This is the system's private notepad of who volunteered for what.
Editing it can scramble assignments.

### Tab 4: "Cancellations" — DO NOT TOUCH

Also the system's private record-keeping.

---

## What Happens Automatically (no action needed)

Every morning the system:

- **6 AM** — assigns docents to tours coming up in the next 14 days
  (never more than one tour per docent per day)
- **7 AM** — emails each docent their reminders and any tours that
  still need help
- **8 AM** — checks if any "Needs Sub" tour went unclaimed too long,
  and emails **you** if a human needs to step in
- **2 AM** — saves a backup copy of the whole spreadsheet

You will get an email whenever the system needs your help — for example,
if nobody can cover a tour. If you don't get an email, everything is fine.

---

## Questions or something looks wrong?

Just email or call Saranya — don't try to fix the Signups or
Cancellations tabs yourself.
