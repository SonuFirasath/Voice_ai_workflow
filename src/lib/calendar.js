// src/lib/calendar.js
const { getAccessToken } = require('./auth');

// The specific Microsoft 365 email address attached to the target calendar
const TARGET_EMAIL = "FirasathG@chimeratechnologies.com"; 
const TIME_ZONE = "Asia/Kolkata";

async function checkAvailability(dateString) {
    console.log(`[CALENDAR] Checking availability for ${dateString}`);
    
    try {
        const accessToken = await getAccessToken();
        if (!accessToken) return { error: "Authentication failed." };

        const endpoint = `https://graph.microsoft.com/v1.0/users/${TARGET_EMAIL}/calendar/getSchedule`;
        
        // Define working hours for the specific date
        const requestBody = {
            schedules: [TARGET_EMAIL],
            startTime: { dateTime: `${dateString}T09:00:00`, timeZone: TIME_ZONE },
            endTime: { dateTime: `${dateString}T17:00:00`, timeZone: TIME_ZONE },
            availabilityViewInterval: 30
        };

        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
                'Prefer': `outlook.timezone="${TIME_ZONE}"`
            },
            body: JSON.stringify(requestBody)
        });

        const data = await response.json();
        
        if (data.error) {
            console.error("[CALENDAR] Graph Error:", data.error);
            return { error: "Failed to access the calendar schedule." };
        }

        const scheduleItems = data.value[0].scheduleItems;
        const busyTimes = scheduleItems.map(item => item.start.dateTime);
        
        return { text: `The calendar has busy slots at these times: ${busyTimes.join(', ')}. Please find an open 30-minute window around these.` };

    } catch (error) {
        return { error: error.message };
    }
}

async function bookAppointment(timeString, topic) {
    console.log(`[CALENDAR] Booking appointment for ${timeString} regarding ${topic}`);
    
    try {
        const accessToken = await getAccessToken();
        if (!accessToken) return { error: "Authentication failed." };

        const endpoint = `https://graph.microsoft.com/v1.0/users/${TARGET_EMAIL}/events`;
        
        // Calculate a strict 30-minute end time
        const startDate = new Date(timeString);
        const endDate = new Date(startDate.getTime() + 30 * 60000);

        const requestBody = {
            subject: `Meeting: ${topic}`,
            start: { dateTime: startDate.toISOString(), timeZone: TIME_ZONE },
            end: { dateTime: endDate.toISOString(), timeZone: TIME_ZONE }
        };

        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestBody)
        });

        const data = await response.json();

        if (data.error) {
            return { error: "Failed to create the calendar event." };
        }

        return { text: "The meeting has been successfully booked on the calendar." };

    } catch (error) {
        return { error: error.message };
    }
}

module.exports = { checkAvailability, bookAppointment };