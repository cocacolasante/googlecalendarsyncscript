/**
 * Automatically synchronizes 'Busy' status from a personal calendar (Source)
 * to a business calendar (Target).
 */

// ====================================================================
// --- CONFIGURATION ---
// ====================================================================
const SOURCE_CALENDAR_ID = 'source@gmail.com';
const TARGET_CALENDAR_ID = 'target@gmail.com';
const SYNC_LOOK_AHEAD_DAYS = 60;
const BUSY_BLOCK_TITLE = "Personal Time Block";
const SYNC_TAG = "SYNC_BLOCK_BY_SCRIPT_V1"; 

/**
 * Main function to be run by the time-driven trigger.
 */
function syncCalendars() {
  try {
    const sourceCalendar = CalendarApp.getCalendarById(SOURCE_CALENDAR_ID);
    const targetCalendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);

    if (!sourceCalendar || !targetCalendar) {
      Logger.log("Error: One or both calendar IDs are invalid. Please check configuration.");
      return;
    }

    const startTime = new Date();
    const endTime = new Date();
    endTime.setDate(endTime.getDate() + SYNC_LOOK_AHEAD_DAYS);
    
    // 1. Get all existing events on the target calendar and filter for our sync blocks
    const allTargetEvents = targetCalendar.getEvents(startTime, endTime);
    const existingTargetBlocks = [];
    
    for (let i = 0; i < allTargetEvents.length; i++) {
      const event = allTargetEvents[i];
      const description = event.getDescription();
      if (description && description.includes(SYNC_TAG)) {
        existingTargetBlocks.push(event);
      }
    }
    
    // Create a map based on start time + end time combination (more reliable than event IDs)
    const timeKeyToTargetBlockMap = new Map();

    // Map existing blocks by their time signature
    for (let i = 0; i < existingTargetBlocks.length; i++) {
      const block = existingTargetBlocks[i];
      const timeKey = block.getStartTime().getTime() + '_' + block.getEndTime().getTime();
      
      // If there's already a block with this time, keep the first one and delete duplicates
      if (timeKeyToTargetBlockMap.has(timeKey)) {
        Logger.log(`Deleting duplicate block at ${block.getStartTime()}`);
        block.deleteEvent();
      } else {
        timeKeyToTargetBlockMap.set(timeKey, block);
      }
    }
    
    // 2. Get events from the source (personal) calendar.
    const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
    const processedTimeKeys = new Set();
    
    // 3. Process each event from the source calendar.
    for (let i = 0; i < sourceEvents.length; i++) {
      const sourceEvent = sourceEvents[i];
      const sourceEventId = sourceEvent.getId();
      const timeKey = sourceEvent.getStartTime().getTime() + '_' + sourceEvent.getEndTime().getTime();
      
      processedTimeKeys.add(timeKey);

      const existingBlock = timeKeyToTargetBlockMap.get(timeKey);

      if (existingBlock) {
        // Block already exists for this time slot
        Logger.log(`Block already exists for time slot: ${sourceEvent.getStartTime()}`);
        
        // Update the description to ensure it has the latest source event ID
        const description = `${BUSY_BLOCK_TITLE} - ${SYNC_TAG}\nSource Event ID: ${sourceEventId}`;
        existingBlock.setDescription(description);
        
        // Remove from map so it won't be deleted later
        timeKeyToTargetBlockMap.delete(timeKey);

      } else {
        // Block does not exist, create a new one.
        const description = `${BUSY_BLOCK_TITLE} - ${SYNC_TAG}\nSource Event ID: ${sourceEventId}`;
        
        // Create event with minimal options
        const newEvent = targetCalendar.createEvent(
          BUSY_BLOCK_TITLE,
          sourceEvent.getStartTime(),
          sourceEvent.getEndTime()
        );
        
        // Set additional properties after creation
        newEvent.setDescription(description);
        newEvent.setVisibility(CalendarApp.Visibility.PRIVATE);
        
        Logger.log(`Created new busy block for source event: ${sourceEvent.getTitle()} at ${sourceEvent.getStartTime()}`);
      }
    }
    
    // 4. Delete any remaining target blocks whose time slots no longer have corresponding source events.
    timeKeyToTargetBlockMap.forEach(function(block, timeKey) {
      Logger.log(`Deleting expired busy block at: ${block.getStartTime()}`);
      block.deleteEvent();
    });

    Logger.log("Calendar sync complete. Processed " + sourceEvents.length + " source events.");

  } catch (e) {
    Logger.log("An error occurred during synchronization: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
  }
}

/**
 * Creates the time-driven trigger for the script to run automatically.
 * Run this function once manually after setting up the script.
 */
function createTrigger() {
  ScriptApp.newTrigger('syncCalendars')
      .timeBased()
      .everyMinutes(5)
      .create();
  Logger.log("Synchronization trigger created successfully. Script will run every 5 minutes.");
}

/**
 * Deletes all triggers associated with this script.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log("All synchronization triggers deleted.");
}
