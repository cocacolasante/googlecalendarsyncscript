/**
 * Automatically synchronizes 'Busy' status from a personal calendar (Source)
 * to a business calendar (Target).
 * QUOTA-OPTIMIZED VERSION
 */

// ====================================================================
// --- CONFIGURATION ---
// ====================================================================
const SOURCE_CALENDAR_ID = 'source@email.com';
const TARGET_CALENDAR_ID = 'target@email.com';
const SYNC_LOOK_AHEAD_DAYS = 14;  
const BUSY_BLOCK_TITLE = "Personal Time Block";

/**
 * Main function to be run by the time-driven trigger.
 */
function syncCalendars() {
  try {
    const sourceCalendar = CalendarApp.getCalendarById(SOURCE_CALENDAR_ID);
    const targetCalendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);

    if (!sourceCalendar || !targetCalendar) {
      Logger.log("Error: One or both calendar IDs are invalid.");
      return;
    }

    // Start from beginning of today (midnight)
    const startTime = new Date();
    startTime.setHours(0, 0, 0, 0);
    
    // End at the last moment of the day X days from now
    const endTime = new Date();
    endTime.setDate(endTime.getDate() + SYNC_LOOK_AHEAD_DAYS);
    endTime.setHours(23, 59, 59, 999);
    
    Logger.log(`Syncing from ${startTime.toDateString()} to ${endTime.toDateString()}`);
    
    // Get events from both calendars
    const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
    const targetEvents = targetCalendar.getEvents(startTime, endTime);
    
    Logger.log(`Found ${sourceEvents.length} events on personal calendar`);
    
    // Filter target events to only our sync blocks
    const existingBlocks = [];
    for (let i = 0; i < targetEvents.length; i++) {
      if (targetEvents[i].getTitle() === BUSY_BLOCK_TITLE) {
        existingBlocks.push(targetEvents[i]);
      }
    }
    
    Logger.log(`Found ${existingBlocks.length} existing blocks on business calendar`);
    
    // Create time-based map of existing blocks
    const existingBlockTimes = new Map();
    for (let i = 0; i < existingBlocks.length; i++) {
      const block = existingBlocks[i];
      const timeKey = block.getStartTime().getTime() + '_' + block.getEndTime().getTime();
      
      if (existingBlockTimes.has(timeKey)) {
        // Delete duplicate immediately
        block.deleteEvent();
        Logger.log(`Removed duplicate at ${block.getStartTime().toLocaleString()}`);
      } else {
        existingBlockTimes.set(timeKey, block);
      }
    }
    
    // Create time-based set of source events
    const sourceEventTimes = new Set();
    for (let i = 0; i < sourceEvents.length; i++) {
      const timeKey = sourceEvents[i].getStartTime().getTime() + '_' + sourceEvents[i].getEndTime().getTime();
      sourceEventTimes.add(timeKey);
    }
    
    // Create missing blocks
    let created = 0;
    for (let i = 0; i < sourceEvents.length; i++) {
      const sourceEvent = sourceEvents[i];
      const timeKey = sourceEvent.getStartTime().getTime() + '_' + sourceEvent.getEndTime().getTime();
      
      if (!existingBlockTimes.has(timeKey)) {
        const newEvent = targetCalendar.createEvent(
          BUSY_BLOCK_TITLE,
          sourceEvent.getStartTime(),
          sourceEvent.getEndTime()
        );
        newEvent.setVisibility(CalendarApp.Visibility.PRIVATE);
        created++;
        Logger.log(`Created block for ${sourceEvent.getTitle()} on ${sourceEvent.getStartTime().toLocaleString()}`);
      } else {
        Logger.log(`Block already exists for ${sourceEvent.getStartTime().toLocaleString()}`);
      }
    }
    
    // Delete expired blocks
    let deleted = 0;
    existingBlockTimes.forEach(function(block, timeKey) {
      if (!sourceEventTimes.has(timeKey)) {
        block.deleteEvent();
        deleted++;
        Logger.log(`Deleted expired block at ${block.getStartTime().toLocaleString()}`);
      }
    });

    Logger.log(`=== SYNC COMPLETE ===`);
    Logger.log(`Created: ${created} new blocks`);
    Logger.log(`Deleted: ${deleted} expired blocks`);
    Logger.log(`Total personal events: ${sourceEvents.length}`);
    Logger.log(`Total business blocks after sync: ${sourceEvents.length}`);

  } catch (e) {
    Logger.log("Error: " + e.toString());
    Logger.log("Stack: " + e.stack);
  }
}

/**
 * Test function to check what events exist
 */
function testCheckEvents() {
  const sourceCalendar = CalendarApp.getCalendarById(SOURCE_CALENDAR_ID);
  const targetCalendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
  
  const startTime = new Date();
  startTime.setHours(0, 0, 0, 0);
  
  const endTime = new Date();
  endTime.setDate(endTime.getDate() + 7);  // Check next 7 days
  endTime.setHours(23, 59, 59, 999);
  
  Logger.log(`\n=== CHECKING EVENTS FROM ${startTime.toDateString()} TO ${endTime.toDateString()} ===\n`);
  
  const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
  Logger.log(`\n--- PERSONAL CALENDAR (${sourceEvents.length} events) ---`);
  for (let i = 0; i < sourceEvents.length; i++) {
    const event = sourceEvents[i];
    Logger.log(`${event.getStartTime().toLocaleString()} - ${event.getTitle()}`);
  }
  
  const targetEvents = targetCalendar.getEvents(startTime, endTime);
  const blocks = [];
  for (let i = 0; i < targetEvents.length; i++) {
    if (targetEvents[i].getTitle() === BUSY_BLOCK_TITLE) {
      blocks.push(targetEvents[i]);
    }
  }
  
  Logger.log(`\n--- BUSINESS CALENDAR (${blocks.length} blocks) ---`);
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    Logger.log(`${block.getStartTime().toLocaleString()} - ${block.getTitle()}`);
  }
  
  Logger.log(`\n=== SUMMARY ===`);
  Logger.log(`Personal events: ${sourceEvents.length}`);
  Logger.log(`Business blocks: ${blocks.length}`);
  Logger.log(`Missing blocks: ${sourceEvents.length - blocks.length}`);
}

/**
 * Creates a time-driven trigger - runs every 30 minutes
 */
function createTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Create new trigger - every 30 minutes
  ScriptApp.newTrigger('syncCalendars')
      .timeBased()
      .everyMinutes(30)
      .create();
  Logger.log("Trigger created: runs every 30 minutes");
}

/**
 * Deletes all triggers
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log("All triggers deleted.");
}
