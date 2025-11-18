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
    
    const sourceIdToTargetIdMap = new Map();

    // Map existing blocks by their source event ID
    for (let i = 0; i < existingTargetBlocks.length; i++) {
      const block = existingTargetBlocks[i];
      const description = block.getDescription();
      if (description) {
        const sourceIdMatch = description.match(/Source Event ID: (.+)/);
        if (sourceIdMatch && sourceIdMatch[1]) {
          sourceIdToTargetIdMap.set(sourceIdMatch[1], block);
        } else {
          // Delete blocks without proper source ID
          block.deleteEvent();
        }
      }
    }
    
    // 2. Get events from the source (personal) calendar.
    const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
    
    // 3. Process each event from the source calendar.
    for (let i = 0; i < sourceEvents.length; i++) {
      const sourceEvent = sourceEvents[i];
      const sourceEventId = sourceEvent.getId();

      const existingBlock = sourceIdToTargetIdMap.get(sourceEventId);

      if (existingBlock) {
        // Block exists, check if times need updating.
        if (existingBlock.getStartTime().getTime() !== sourceEvent.getStartTime().getTime() ||
            existingBlock.getEndTime().getTime() !== sourceEvent.getEndTime().getTime()) {
          
          existingBlock.setTime(sourceEvent.getStartTime(), sourceEvent.getEndTime());
          Logger.log(`Updated busy block for source event: ${sourceEvent.getTitle()}`);
        }
        sourceIdToTargetIdMap.delete(sourceEventId); 

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
        
        Logger.log(`Created new busy block for source event: ${sourceEvent.getTitle()}`);
      }
    }
    
    // 4. Delete any remaining target blocks whose corresponding source event no longer exists.
    sourceIdToTargetIdMap.forEach(function(block) {
      Logger.log(`Deleting expired busy block: ${block.getTitle()}`);
      block.deleteEvent();
    });

    Logger.log("Calendar sync complete.");

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
