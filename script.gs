/**
 * Automatically synchronizes 'Busy' status from a personal calendar (Source)
 * to a business calendar (Target).
 * QUOTA-OPTIMIZED VERSION
 */

// ====================================================================
// --- CONFIGURATION ---
// ====================================================================
const SOURCE_CALENDAR_ID = 'source@gmail.com';
const TARGET_CALENDAR_ID = 'source@gmail.com';
const SYNC_LOOK_AHEAD_DAYS = 14;  // Reduced from 60 to save quota
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

    const now = new Date();
    const startTime = new Date(now.getTime());
    const endTime = new Date(now.getTime());
    endTime.setDate(endTime.getDate() + SYNC_LOOK_AHEAD_DAYS);
    
    Logger.log(`Syncing from ${startTime.toDateString()} to ${endTime.toDateString()}`);
    
    // Get events from both calendars
    const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
    const targetEvents = targetCalendar.getEvents(startTime, endTime);
    
    // Filter target events to only our sync blocks
    const existingBlocks = [];
    for (let i = 0; i < targetEvents.length; i++) {
      if (targetEvents[i].getTitle() === BUSY_BLOCK_TITLE) {
        existingBlocks.push(targetEvents[i]);
      }
    }
    
    Logger.log(`Source events: ${sourceEvents.length}, Existing blocks: ${existingBlocks.length}`);
    
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
        targetCalendar.createEvent(
          BUSY_BLOCK_TITLE,
          sourceEvent.getStartTime(),
          sourceEvent.getEndTime()
        );
        created++;
        Logger.log(`Created block for ${sourceEvent.getStartTime().toLocaleString()}`);
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

    Logger.log(`Sync complete. Created: ${created}, Deleted: ${deleted}`);

  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
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
 * One-time cleanup function to remove all existing duplicate blocks.
 * Run this manually once, then delete it or comment it out.
 */
function cleanupAllDuplicates() {
  try {
    const targetCalendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
    
    const startTime = new Date();
    const endTime = new Date();
    endTime.setDate(endTime.getDate() + SYNC_LOOK_AHEAD_DAYS);
    
    const allEvents = targetCalendar.getEvents(startTime, endTime);
    const timeKeyMap = new Map();
    let deletedCount = 0;
    
    for (let i = 0; i < allEvents.length; i++) {
      const event = allEvents[i];
      if (event.getTitle() === BUSY_BLOCK_TITLE) {
        const timeKey = event.getStartTime().getTime() + '_' + event.getEndTime().getTime();
        
        if (timeKeyMap.has(timeKey)) {
          Logger.log(`Deleting duplicate at ${event.getStartTime()}`);
          event.deleteEvent();
          deletedCount++;
        } else {
          timeKeyMap.set(timeKey, event);
        }
      }
    }
    
    Logger.log(`Cleanup complete. Deleted ${deletedCount} duplicate blocks.`);
  } catch (e) {
    Logger.log("Cleanup error: " + e.toString());
  }
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
