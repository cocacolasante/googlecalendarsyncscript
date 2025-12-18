/**
 * Automatically synchronizes 'Busy' status from a personal calendar (Source)
 * to a business calendar (Target).
 * QUOTA-OPTIMIZED VERSION
 */

// ====================================================================
// --- CONFIGURATION ---
// ====================================================================
const SOURCE_CALENDAR_ID = 'source@gmail.com';
const TARGET_CALENDAR_ID = 'target@csuitecode.com';
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
    
    // FIXED: Create time-based map of existing blocks and only delete TRUE duplicates
    const existingBlockTimes = new Map();
    const duplicatesToDelete = [];
    
    for (let i = 0; i < existingBlocks.length; i++) {
      const block = existingBlocks[i];
      const timeKey = block.getStartTime().getTime() + '_' + block.getEndTime().getTime();
      
      if (existingBlockTimes.has(timeKey)) {
        // This is a TRUE duplicate - mark for deletion
        duplicatesToDelete.push(block);
        Logger.log(`Found duplicate at ${block.getStartTime().toLocaleString()}`);
      } else {
        // First block at this time - keep it in the map
        existingBlockTimes.set(timeKey, block);
      }
    }
    
    // Delete only the TRUE duplicates
    for (let i = 0; i < duplicatesToDelete.length; i++) {
      duplicatesToDelete[i].deleteEvent();
      Logger.log(`Removed duplicate at ${duplicatesToDelete[i].getStartTime().toLocaleString()}`);
    }
    
    // Create time-based set of source events with more lenient matching
    const sourceEventTimes = new Set();
    for (let i = 0; i < sourceEvents.length; i++) {
      const sourceEvent = sourceEvents[i];
      const timeKey = sourceEvent.getStartTime().getTime() + '_' + sourceEvent.getEndTime().getTime();
      sourceEventTimes.add(timeKey);
      Logger.log(`Source event: ${sourceEvent.getTitle()} at ${sourceEvent.getStartTime().toLocaleString()}`);
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
        Logger.log(`✓ Created block for ${sourceEvent.getTitle()} on ${sourceEvent.getStartTime().toLocaleString()}`);
      } else {
        Logger.log(`✓ Block already exists for ${sourceEvent.getStartTime().toLocaleString()}`);
      }
    }
    
    // FIXED: Only delete blocks that are truly expired (no matching source event)
    // BUT be more careful about timing comparisons
    let deleted = 0;
    const blocksToDelete = [];
    
    existingBlockTimes.forEach(function(block, timeKey) {
      if (!sourceEventTimes.has(timeKey)) {
        // Double-check: is there a source event at this exact time?
        let foundMatch = false;
        for (let i = 0; i < sourceEvents.length; i++) {
          const sourceEvent = sourceEvents[i];
          if (sourceEvent.getStartTime().getTime() === block.getStartTime().getTime() &&
              sourceEvent.getEndTime().getTime() === block.getEndTime().getTime()) {
            foundMatch = true;
            Logger.log(`⚠ Block at ${block.getStartTime().toLocaleString()} has matching source event - keeping it`);
            break;
          }
        }
        
        if (!foundMatch) {
          blocksToDelete.push(block);
          Logger.log(`✗ Marking block for deletion at ${block.getStartTime().toLocaleString()} - no matching source event found`);
        }
      }
    });
    
    // Only delete after confirming
    for (let i = 0; i < blocksToDelete.length; i++) {
      blocksToDelete[i].deleteEvent();
      deleted++;
      Logger.log(`Deleted expired block at ${blocksToDelete[i].getStartTime().toLocaleString()}`);
    }

    Logger.log(`\n=== SYNC COMPLETE ===`);
    Logger.log(`Created: ${created} new blocks`);
    Logger.log(`Deleted: ${deleted} expired blocks`);
    Logger.log(`Removed: ${duplicatesToDelete.length} duplicates`);
    Logger.log(`Total personal events: ${sourceEvents.length}`);
    Logger.log(`Expected business blocks: ${sourceEvents.length}`);

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
  endTime.setDate(endTime.getDate() + 7);
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
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
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
