/**
 * @OnlyCurrentDoc
 *
 * สคริปต์ฝั่งเซิร์ฟเวอร์สำหรับระบบจองห้องประชุมและรถยนต์
 * * MODIFIED: เปลี่ยนระบบให้ Run as User Accessing (ผู้จองเป็น Organizer)
 * * MODIFIED: อัปเดตเวอร์ชันเพื่อ Force Clear Cache
 */

// --- การตั้งค่าเริ่มต้น ---
// *** อย่าลืมตรวจสอบ ID ของไฟล์จริงของคุณ ***
const SPREADSHEET_ID = "1RVzYj-D098P_Cgq6t45eG9uigJqRiEpC3wTjspWdEq4"; // <--- ตรวจสอบ ID
const APP_VERSION = "7.2-ForceClear"; // <--- เปลี่ยนเลขตรงนี้เมื่อต้องการบังคับ User เคลียร์ Cache

// Sheet Names
const SHEET_NAME_ROOMS = "Rooms";
const SHEET_NAME_ROOM_BOOKINGS = "Bookings";
const SHEET_NAME_CARS = "Cars";
const SHEET_NAME_CAR_BOOKINGS = "CarBookings";
const SHEET_NAME_ADMINS = "Admins";

// Standard Headers (เพิ่ม IsPrivate)
const ROOM_BOOKING_HEADERS = ["BookingID", "Timestamp", "Title", "Room", "StartTime", "EndTime", "BookedBy", "Status", "CancelledBy", "CancelledTimestamp", "Attendees", "CalendarEventId", "MeetLink", "IsPrivate"];
const CAR_BOOKING_HEADERS = ["BookingID", "Timestamp", "Title", "Car", "StartTime", "EndTime", "BookedBy", "Status", "CancelledBy", "CancelledTimestamp", "Attendees", "CalendarEventId", "IsPrivate"];


// --- Main Functions ---

function doGet(e) {
  const userIsAdmin = checkIfUserIsAdmin_();
  
  let template;
  if (e.parameter.page === 'admin' && userIsAdmin) {
    template = HtmlService.createTemplateFromFile('Admin');
  } else {
    template = HtmlService.createTemplateFromFile('WebApp');
  }
  
  template.appVersion = APP_VERSION;

  return template.evaluate()
    .setTitle('BeNeat Central Reservation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  const user = Session.getActiveUser();
  const timezone = Session.getScriptTimeZone();
  const webAppUrl = ScriptApp.getService().getUrl(); 
  
  return {
    userEmail: user ? user.getEmail() : '',
    isAdmin: checkIfUserIsAdmin_(), 
    adminUrl: `${webAppUrl}?page=admin`,
    rooms: getResources_(SHEET_NAME_ROOMS, "RoomName", "ห้องประชุมใหญ่"),
    cars: getResources_(SHEET_NAME_CARS, "CarName", "Toyota Vios"),
    roomBookings: getAllBookings_(SHEET_NAME_ROOM_BOOKINGS, timezone),
    carBookings: getAllBookings_(SHEET_NAME_CAR_BOOKINGS, timezone),
    directory: getWorkspaceUsers()
  };
}

function getAdminDashboardData() {
  checkIfUserIsAdminOrThrow_(); 
  
  const timezone = Session.getScriptTimeZone();
  const user = Session.getActiveUser();
  
  return {
    userEmail: user ? user.getEmail() : '',
    rooms: getResources_(SHEET_NAME_ROOMS, "RoomName", null),
    cars: getResources_(SHEET_NAME_CARS, "CarName", null),
    admins: getAdmins(),
    roomBookings: getAllBookingsForAdmin_(SHEET_NAME_ROOM_BOOKINGS, timezone, "Room"),
    carBookings: getAllBookingsForAdmin_(SHEET_NAME_CAR_BOOKINGS, timezone, "Car"),
  };
}

/**
 * ดึงรายชื่อผู้ใช้
 */
function getWorkspaceUsers() {
  if (typeof AdminDirectory === 'undefined') {
    return [];
  }

  try {
    const users = [];
    let pageToken;
    do {
      const response = AdminDirectory.Users.list({
        customer: 'my_customer', 
        maxResults: 500,
        projection: "basic",
        viewType: 'domain_public', 
        pageToken: pageToken
      });
      
      const pageUsers = response.users;
      if (pageUsers && pageUsers.length > 0) {
        for (let i = 0; i < pageUsers.length; i++) {
          const user = pageUsers[i];
          if (!user.suspended) {
            users.push({
              name: user.name.fullName,
              email: user.primaryEmail
            });
          }
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    return users;
  } catch (e) {
    Logger.log("User cannot access Admin Directory: " + e.message);
    return []; 
  }
}

function checkAttendeesAvailability(emails, startTimeStr, endTimeStr) {
  try {
    const start = new Date(startTimeStr);
    const end = new Date(endTimeStr);
    const busyEmails = [];

    emails.forEach(email => {
      try {
        const cal = CalendarApp.getCalendarById(email.trim());
        if (cal) {
          const events = cal.getEvents(start, end);
          if (events.length > 0) {
            busyEmails.push(email.trim());
          }
        }
      } catch (err) {
      }
    });

    return { success: true, busyEmails: busyEmails };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- Generic Resource Management ---

function getResources_(sheetName, header, defaultItem) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([header]);
      if (defaultItem) {
          sheet.appendRow([defaultItem]);
          return [defaultItem];
      }
      return [];
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  } catch (e) {
    Logger.log(`Error getting resources from ${sheetName}: ` + e.message);
    return [];
  }
}

function addRoom(roomName) {
  checkIfUserIsAdminOrThrow_(); 
  return addResource_(SHEET_NAME_ROOMS, "RoomName", roomName, "ห้องประชุม");
}

function addCar(carName) {
  checkIfUserIsAdminOrThrow_(); 
  return addResource_(SHEET_NAME_CARS, "CarName", carName, "รถยนต์");
}

function addResource_(sheetName, header, resourceName, resourceType) {
  if (!resourceName || resourceName.trim() === "") {
    throw new Error(`ชื่อ${resourceType}ห้ามว่าง`);
  }
  const trimmedName = resourceName.trim();
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found.`);
    }
    const lastRow = sheet.getLastRow();
    const existingItems = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(r => String(r).toLowerCase().trim()) : [];
    
    if (existingItems.includes(trimmedName.toLowerCase())) {
      throw new Error(`${resourceType}ชื่อ '${trimmedName}' มีอยู่แล้ว`);
    }
    sheet.appendRow([trimmedName]);
    return { success: true, message: `เพิ่ม${resourceType}สำเร็จ`, newResource: trimmedName }; 
  } catch (e) {
    throw new Error(e.message || `เกิดข้อผิดพลาดในการเพิ่ม${resourceType}`);
  }
}

function deleteRoom(roomName) {
  checkIfUserIsAdminOrThrow_(); 
  return deleteResource_(SHEET_NAME_ROOMS, SHEET_NAME_ROOM_BOOKINGS, "Room", roomName, "ห้องประชุม");
}

function deleteCar(carName) {
  checkIfUserIsAdminOrThrow_(); 
  return deleteResource_(SHEET_NAME_CARS, SHEET_NAME_CAR_BOOKINGS, "Car", carName, "รถยนต์");
}

function deleteResource_(resourceSheetName, bookingSheetName, resourceHeader, resourceName, resourceType) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const bookingSheet = ss.getSheetByName(bookingSheetName);
  
  if (bookingSheet && bookingSheet.getLastRow() > 1) {
    const data = bookingSheet.getDataRange().getValues();
    const headers = data.shift();
    const resourceCol = headers.indexOf(resourceHeader);
    const statusCol = headers.indexOf("Status");
    const endCol = headers.indexOf("EndTime");
    
    const now = new Date().getTime();

    const hasActiveBooking = data.some(row => 
      row[resourceCol] === resourceName && 
      row[statusCol] === 'Confirmed' &&
      parseSheetDate_(row[endCol]).getTime() > now 
    );
    
    if (hasActiveBooking) {
      throw new Error(`ไม่สามารถลบ '${resourceName}' ได้ เนื่องจากยังมีการจองที่ Active อยู่`);
    }
  }

  const resourceSheet = ss.getSheetByName(resourceSheetName);
  if (!resourceSheet) throw new Error(`ไม่พบชีต '${resourceSheetName}'`);
  
  const data = resourceSheet.getRange(2, 1, resourceSheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(row => row[0] === resourceName);
  
  if (rowIndex === -1) {
    throw new Error(`ไม่พบ ${resourceType} ชื่อ '${resourceName}'`);
  }
  
  resourceSheet.deleteRow(rowIndex + 2); 
  
  return { success: true, message: `ลบ ${resourceType} '${resourceName}' สำเร็จ`, removedResource: resourceName };
}

function updateRoom(oldName, newName) {
  checkIfUserIsAdminOrThrow_();
  return updateResource_(SHEET_NAME_ROOMS, SHEET_NAME_ROOM_BOOKINGS, "Room", oldName, newName, "ห้องประชุม");
}

function updateCar(oldName, newName) {
  checkIfUserIsAdminOrThrow_();
  return updateResource_(SHEET_NAME_CARS, SHEET_NAME_CAR_BOOKINGS, "Car", oldName, newName, "รถยนต์");
}

function updateResource_(resourceSheetName, bookingSheetName, resourceHeader, oldName, newName, resourceType) {
  const trimmedNewName = newName.trim();
  if (!trimmedNewName) {
    throw new Error(`ชื่อ${resourceType}ใหม่ห้ามว่าง`);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resourceSheet = ss.getSheetByName(resourceSheetName);
  if (!resourceSheet) throw new Error(`ไม่พบชีต '${resourceSheetName}'`);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000); 

  try {
    const lastRow = resourceSheet.getLastRow();
    const existingItems = lastRow > 1 ? resourceSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(r => String(r).toLowerCase().trim()) : [];
    
    if (existingItems.includes(trimmedNewName.toLowerCase())) {
      throw new Error(`${resourceType}ชื่อ '${trimmedNewName}' มีอยู่แล้ว`);
    }

    const data = resourceSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const rowIndex = data.findIndex(row => String(row[0]).trim() === String(oldName).trim());
    
    if (rowIndex === -1) {
      throw new Error(`ไม่พบ ${resourceType} ชื่อเดิม '${oldName}'`);
    }

    resourceSheet.getRange(rowIndex + 2, 1).setValue(trimmedNewName);

    const bookingSheet = ss.getSheetByName(bookingSheetName);
    if (bookingSheet && bookingSheet.getLastRow() > 1) {
      const bookingData = bookingSheet.getDataRange().getValues();
      const headers = bookingData.shift();
      const resourceColIndex = headers.indexOf(resourceHeader);

      if (resourceColIndex !== -1) {
        const rangesToUpdate = [];
        bookingData.forEach((row, index) => {
          if (String(row[resourceColIndex]).trim() === String(oldName).trim()) {
            rangesToUpdate.push(bookingSheet.getRange(index + 2, resourceColIndex + 1));
          }
        });

        if (rangesToUpdate.length > 0) {
          rangesToUpdate.forEach(range => range.setValue(trimmedNewName));
        }
      }
    }
    
    return { 
      success: true, 
      message: `อัปเดต ${resourceType} เป็น '${trimmedNewName}' สำเร็จ`, 
      oldName: oldName, 
      newName: trimmedNewName 
    };

  } catch (e) {
    throw new Error(e.message || `เกิดข้อผิดพลาดในการอัปเดต${resourceType}`);
  } finally {
    lock.releaseLock();
  }
}

// --- Generic Booking Management ---

function submitRoomBooking(formData) {
  return submitBooking_(SHEET_NAME_ROOM_BOOKINGS, "Room", ROOM_BOOKING_HEADERS, formData);
}

function submitCarBooking(formData) {
  return submitBooking_(SHEET_NAME_CAR_BOOKINGS, "Car", CAR_BOOKING_HEADERS, formData);
}

function submitBooking_(sheetName, resourceHeader, headersConst, formData) {
  const userEmail = getUserEmail_(); 
  const { title, resource, startTime, endTime, attendees, isPrivate } = formData; 
  
  if (!title || !resource || !startTime || !endTime) throw new Error("กรุณากรอกข้อมูลให้ครบถ้วน");

  const start = new Date(startTime);
  const end = new Date(endTime);
  if (start >= end) throw new Error("เวลาสิ้นสุดต้องอยู่หลังเวลาเริ่มต้น");

  try {
    checkForConflict_(sheetName, resourceHeader, resource, start, end);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headersConst); 
      ensureHeader_(sheet, headersConst); 
    } else {
      ensureHeader_(sheet, headersConst); 
    }
    
    const bookingId = "BK" + new Date().getTime();
    
    // --- 📅 Create Google Calendar Event & Meet Link ---
    let calendarEventId = "";
    let meetLink = "";
    
    if (resourceHeader === "Room") {
        try {
            const description = `รายละเอียดการจอง: ${resource}\nผู้จอง: ${userEmail}\nระบบจอง BeNeat Central Reservation`;
            const eventTitle = title; 

            let attendeeList = [];
            
            if (attendees && attendees.trim() !== "") {
                const guestArray = attendees.split(',');
                attendeeList = guestArray.map(email => ({ email: email.trim() }));
            }

            const eventPayload = {
                summary: eventTitle,
                location: resource,
                description: description,
                start: { dateTime: start.toISOString() },
                end: { dateTime: end.toISOString() },
                attendees: attendeeList,
                guestsCanModify: true, 
                visibility: isPrivate ? 'private' : 'default',
                conferenceData: {
                    createRequest: {
                        requestId: bookingId, 
                        conferenceSolutionKey: { type: "hangoutsMeet" }
                    }
                }
            };

            const createdEvent = Calendar.Events.insert(eventPayload, 'primary', {
                conferenceDataVersion: 1,
                sendUpdates: "all" 
            });
            
            calendarEventId = createdEvent.id;
            if (createdEvent.hangoutLink) {
                meetLink = createdEvent.hangoutLink;
            }
        } catch (calError) {
            Logger.log("สร้างปฏิทินไม่สำเร็จ: " + calError.message);
        }
    }
    // ----------------------------------------

    const newRowData = {
      BookingID: bookingId, Timestamp: new Date(), Title: title,
      [resourceHeader]: resource, StartTime: start, EndTime: end,
      BookedBy: userEmail, Status: "Confirmed", CancelledBy: "", CancelledTimestamp: "",
      Attendees: attendees || "", CalendarEventId: calendarEventId,
      MeetLink: meetLink,
      IsPrivate: isPrivate ? "TRUE" : "FALSE" 
    };

    const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowArray = actualHeaders.map(header => newRowData[header.trim()] || "");
    
    sheet.appendRow(rowArray);

    const newEvent = {
        title: `${title} (${userEmail.split('@')[0]})`,
        start: start.toISOString(),
        end: end.toISOString(),
        [resourceHeader.toLowerCase()]: resource,
        extendedProps: {
            bookingId: bookingId,
            bookedBy: userEmail,
            fullTitle: title,
            attendees: attendees || "",
            meetLink: meetLink,
            isPrivate: isPrivate 
        }
    };
    
    return { 
      success: true, 
      message: "บันทึกการจองและอัปเดตปฏิทินสำเร็จ!", 
      newBooking: newEvent,
      rooms: getResources_(SHEET_NAME_ROOMS, "RoomName", null),
      cars: getResources_(SHEET_NAME_CARS, "CarName", null),
      roomBookings: getAllBookings_(SHEET_NAME_ROOM_BOOKINGS, Session.getScriptTimeZone()),
      carBookings: getAllBookings_(SHEET_NAME_CAR_BOOKINGS, Session.getScriptTimeZone()),
      isAdmin: checkIfUserIsAdmin_()
    };
  } catch (e) {
    Logger.log(`Error submitting booking to ${sheetName}: ` + e.message);
    throw new Error(e.message || "เกิดข้อผิดพลาดในการบันทึกข้อมูล (กรุณาตรวจสอบสิทธิ์การเข้าถึง Sheet)");
  }
}

function updateRoomBooking(formData) {
  return updateBooking_(SHEET_NAME_ROOM_BOOKINGS, "Room", formData);
}

function updateCarBooking(formData) {
  return updateBooking_(SHEET_NAME_CAR_BOOKINGS, "Car", formData);
}

function updateBooking_(sheetName, resourceHeader, formData) {
    const { bookingId, title, startTime, endTime, attendees, isPrivate } = formData;

    if (!bookingId || !title || !startTime || !endTime) {
        throw new Error("ข้อมูลไม่ครบถ้วนสำหรับการอัปเดต");
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    if (start >= end) {
        throw new Error("เวลาสิ้นสุดต้องอยู่หลังเวลาเริ่มต้น");
    }
    
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
        const { row, rowIndex, headers } = findBookingRowById_(sheet, bookingId);

        // --- ตรวจสอบสิทธิ์ ---
        const bookedByCol = headers.indexOf("BookedBy");
        const userEmail = getUserEmail_();
        const creatorEmail = row[bookedByCol]; 

        if (creatorEmail !== userEmail && !checkIfUserIsAdmin_()) {
           throw new Error("คุณไม่มีสิทธิ์แก้ไขการจองนี้");
        }

        const resourceCol = headers.indexOf(resourceHeader);
        const resourceName = row[resourceCol];

        checkForConflict_(sheetName, resourceHeader, resourceName, start, end, bookingId);

        const titleCol = headers.indexOf("Title") + 1;
        const startCol = headers.indexOf("StartTime") + 1;
        const endCol = headers.indexOf("EndTime") + 1;
        const attendeesCol = headers.indexOf("Attendees") + 1;
        const calendarEventIdCol = headers.indexOf("CalendarEventId"); 
        const meetLinkCol = headers.indexOf("MeetLink");
        const isPrivateCol = headers.indexOf("IsPrivate") + 1;

        sheet.getRange(rowIndex, titleCol).setValue(title);
        sheet.getRange(rowIndex, startCol).setValue(start);
        sheet.getRange(rowIndex, endCol).setValue(end);
        
        if(attendeesCol > 0) {
            sheet.getRange(rowIndex, attendeesCol).setValue(attendees || "");
        }

        if(isPrivateCol > 0) {
            sheet.getRange(rowIndex, isPrivateCol).setValue(isPrivate ? "TRUE" : "FALSE");
        }
        
        // --- 📅 Update Google Calendar Event ---
        if (resourceHeader === "Room" && calendarEventIdCol !== -1 && row[calendarEventIdCol]) {
            try {
                const eventId = row[calendarEventIdCol];
                const event = Calendar.Events.get('primary', eventId);
                
                if (event) {
                    const updatePayload = {
                        summary: title, 
                        start: { dateTime: start.toISOString() },
                        end: { dateTime: end.toISOString() },
                        description: event.description,
                        visibility: isPrivate ? 'private' : 'default' 
                    };

                    if (attendees !== undefined) {
                        let newAttendees = [];
                        if (attendees && attendees.trim() !== "") {
                             newAttendees = attendees.split(',').map(email => ({ email: email.trim() }));
                        }
                        updatePayload.attendees = newAttendees;
                    }
                    
                    Calendar.Events.patch(updatePayload, 'primary', eventId);
                }
            } catch (calError) {
                Logger.log("ไม่สามารถอัปเดตปฏิทินได้: " + calError.message);
            }
        }
        // ----------------------------------------
        
        let existingMeetLink = "";
        if (meetLinkCol !== -1 && row[meetLinkCol]) {
            existingMeetLink = row[meetLinkCol];
        }

        const updatedEvent = {
            title: `${title} (${creatorEmail.split('@')[0]})`,
            start: start.toISOString(),
            end: end.toISOString(),
            [resourceHeader.toLowerCase()]: resourceName,
            extendedProps: {
                bookingId: bookingId,
                bookedBy: creatorEmail,
                fullTitle: title,
                attendees: attendees || "",
                meetLink: existingMeetLink,
                isPrivate: isPrivate
            }
        };
        
        return { 
          success: true, 
          message: "อัปเดตการจองสำเร็จ", 
          updatedBooking: updatedEvent,
          rooms: getResources_(SHEET_NAME_ROOMS, "RoomName", null),
          cars: getResources_(SHEET_NAME_CARS, "CarName", null),
          roomBookings: getAllBookings_(SHEET_NAME_ROOM_BOOKINGS, Session.getScriptTimeZone()),
          carBookings: getAllBookings_(SHEET_NAME_CAR_BOOKINGS, Session.getScriptTimeZone()),
          isAdmin: checkIfUserIsAdmin_()
        };
    } catch (e) {
        Logger.log(`Error updating booking in ${sheetName}: ` + e.message);
        throw new Error(e.message || "เกิดข้อผิดพลาดในการอัปเดตการจอง");
    }
}

function cancelRoomBooking(bookingId) {
  return cancelBooking_(SHEET_NAME_ROOM_BOOKINGS, ROOM_BOOKING_HEADERS, bookingId);
}

function cancelCarBooking(bookingId) {
  return cancelBooking_(SHEET_NAME_CAR_BOOKINGS, CAR_BOOKING_HEADERS, bookingId);
}

function cancelBooking_(sheetName, headersConst, bookingId) {
    const userEmail = getUserEmail_(); 
    const cancellationTimestamp = new Date(); 
    
    if (!bookingId) throw new Error("ไม่พบ Booking ID");
    
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
        
        ensureHeader_(sheet, headersConst); 

        const { row, rowIndex, headers } = findBookingRowById_(sheet, bookingId);

        const bookedByCol = headers.indexOf("BookedBy");
        if (row[bookedByCol] !== userEmail && !checkIfUserIsAdmin_()) {
           throw new Error("คุณไม่มีสิทธิ์ยกเลิกการจองนี้");
        }

        const statusCol = headers.indexOf("Status") + 1;
        const cancelledByCol = headers.indexOf("CancelledBy") + 1;
        const cancelledTimestampCol = headers.indexOf("CancelledTimestamp") + 1;
        const calendarEventIdCol = headers.indexOf("CalendarEventId"); 

        const timezone = Session.getScriptTimeZone(); 

        if (row[headers.indexOf("Status")] === "Cancelled") {
          return { 
            success: true, 
            message: "การจองนี้ถูกยกเลิกไปแล้ว", 
            cancelledId: bookingId,
            roomBookings: getAllBookingsForAdmin_(SHEET_NAME_ROOM_BOOKINGS, timezone, "Room"),
            carBookings: getAllBookingsForAdmin_(SHEET_NAME_CAR_BOOKINGS, timezone, "Car")
          };
        }

        sheet.getRange(rowIndex, statusCol).setValue("Cancelled");
        
        if (cancelledByCol > 0) {
            sheet.getRange(rowIndex, cancelledByCol).setValue(userEmail);
        }
        
        if (cancelledTimestampCol > 0) { 
            sheet.getRange(rowIndex, cancelledTimestampCol).setValue(cancellationTimestamp);
        }
        
        // --- 📅 Delete Google Calendar Event ---
        if (calendarEventIdCol !== -1 && row[calendarEventIdCol]) {
            try {
                const eventId = row[calendarEventIdCol];
                Calendar.Events.remove('primary', eventId, { sendUpdates: "all" });
            } catch (calError) {
                Logger.log("ไม่สามารถลบปฏิทินได้ (อาจถูกลบไปแล้ว หรือสิทธิ์ไม่พอ): " + calError.message);
            }
        }
        // ----------------------------------------

        return { 
          success: true, 
          message: "ยกเลิกการจองสำเร็จ", 
          cancelledId: bookingId,
          roomBookings: getAllBookingsForAdmin_(SHEET_NAME_ROOM_BOOKINGS, timezone, "Room"),
          carBookings: getAllBookingsForAdmin_(SHEET_NAME_CAR_BOOKINGS, timezone, "Car")
        };

    } catch (e) {
        Logger.log(`Error cancelling booking in ${sheetName}: ` + e.message);
        throw new Error(e.message || "เกิดข้อผิดพลาดในการยกเลิกการจอง");
    }
}

// --- Generic Helper Functions ---

function getAllBookings_(sheetName, timezone) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() < 2) return [];

        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const resourceHeader = headers.includes("Room") ? "Room" : "Car";
        
        // Cache permissions
        const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
        const isAdmin = checkIfUserIsAdmin_();

        const colIndices = {
            BookingID: headers.indexOf("BookingID"), Title: headers.indexOf("Title"),
            Resource: headers.indexOf(resourceHeader), StartTime: headers.indexOf("StartTime"),
            EndTime: headers.indexOf("EndTime"), Status: headers.indexOf("Status"),
            BookedBy: headers.indexOf("BookedBy"), Attendees: headers.indexOf("Attendees"),
            MeetLink: headers.indexOf("MeetLink"), IsPrivate: headers.indexOf("IsPrivate")
        };

        if (colIndices.BookingID === -1 || colIndices.Resource === -1 || colIndices.StartTime === -1 || colIndices.EndTime === -1 || colIndices.Status === -1 || colIndices.BookedBy === -1) {
            return [];
        }

        return data.filter(row => row[colIndices.Status] === 'Confirmed').map((row, index) => {
            try {
                const start = parseSheetDate_(row[colIndices.StartTime]); 
                const end = parseSheetDate_(row[colIndices.EndTime]); 
                
                if (!start || !end) {
                    return null;
                }

                const bookedBy = row[colIndices.BookedBy];
                const isOwner = bookedBy.toLowerCase() === currentUserEmail;
                // --- FIX: Robust boolean check ---
                // รองรับทั้ง "TRUE" (String), "true" (String), หรือ true (Boolean)
                const isPrivateVal = colIndices.IsPrivate !== -1 ? row[colIndices.IsPrivate] : false;
                const isPrivate = String(isPrivateVal).toUpperCase() === "TRUE";

                // --- 🔒 Privacy Logic ---
                let displayTitle = row[colIndices.Title];
                let displayAttendees = colIndices.Attendees !== -1 ? row[colIndices.Attendees] : "";

                // ถ้าเป็น Private และไม่ใช่เจ้าของ และไม่ใช่ Admin -> ซ่อนข้อมูล
                if (isPrivate && !isOwner && !isAdmin) {
                    displayTitle = "🔒 Private Meeting";
                    displayAttendees = ""; // ซ่อนผู้เข้าร่วมด้วยเพื่อความปลอดภัย
                }
                // -----------------------

                return {
                    title: `${displayTitle} (${bookedBy.split('@')[0]})`, 
                    start: start.toISOString(), end: end.toISOString(),
                    [resourceHeader.toLowerCase()]: row[colIndices.Resource],
                    extendedProps: {
                        bookingId: row[colIndices.BookingID],
                        bookedBy: bookedBy,
                        fullTitle: displayTitle, 
                        attendees: displayAttendees,
                        meetLink: colIndices.MeetLink !== -1 ? row[colIndices.MeetLink] : "",
                        isPrivate: isPrivate
                    }
                };
            } catch(e) {
                return null;
            }
        }).filter(Boolean); 
    } catch (e) {
        return [];
    }
}

function getAllBookingsForAdmin_(sheetName, timezone, resourceHeader) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).trim()); 

    const colIndices = {};
    headers.forEach((header, index) => {
        colIndices[header] = index;
    });

    const needed = ["BookingID", resourceHeader, "StartTime", "EndTime", "Status"];
    if (!needed.every(h => headers.includes(h))) return [];
    
    const type = resourceHeader.toLowerCase(); 

    return data.map((row, index) => {
      try {
        const start = parseSheetDate_(row[colIndices.StartTime]); 
        const end = parseSheetDate_(row[colIndices.EndTime]); 
        
        if (!start || !end) {
            return null;
        }

        const timestamp = row[colIndices.Timestamp] ? parseSheetDate_(row[colIndices.Timestamp]).toISOString() : '';
        const cancelledBy = row[colIndices.CancelledBy] || '';
        const cancelledTimestamp = row[colIndices.CancelledTimestamp] ? parseSheetDate_(row[colIndices.CancelledTimestamp]).toISOString() : '';

        return {
          BookingID: row[colIndices.BookingID],
          Title: row[colIndices.Title] || '',
          resourceName: row[colIndices[resourceHeader]],
          StartTime: start.toISOString(),
          EndTime: end.toISOString(),
          BookedBy: row[colIndices.BookedBy],
          Status: row[colIndices.Status],
          Timestamp: timestamp, 
          CancelledBy: cancelledBy, 
          CancelledTimestamp: cancelledTimestamp, 
          Attendees: colIndices.Attendees !== undefined && colIndices.Attendees !== -1 ? row[colIndices.Attendees] : '',
          type: type
        };
      } catch(e) {
         return null;
      }
    }).filter(Boolean);

  } catch (e) {
    return []; 
  }
}

function checkForConflict_(sheetName, resourceHeader, resource, start, end, existingBookingId = null) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const colIndices = {
    BookingID: headers.indexOf("BookingID"), Resource: headers.indexOf(resourceHeader),
    StartTime: headers.indexOf("StartTime"), EndTime: headers.indexOf("EndTime"),
    Status: headers.indexOf("Status")
  };
  
  const newStart = new Date(start).getTime();
  const newEnd = new Date(end).getTime();

  for (const row of data) {
    const bookingId = row[colIndices.BookingID];
    const existingResource = row[colIndices.Resource];
    const status = row[colIndices.Status];

    if (existingResource === resource && status === 'Confirmed' && bookingId !== existingBookingId) {
      const existingStart = parseSheetDate_(row[colIndices.StartTime]); 
      const existingEnd = parseSheetDate_(row[colIndices.EndTime]); 
      
      if(existingStart && existingEnd) {
          if (newStart < existingEnd.getTime() && newEnd > existingStart.getTime()) {
            throw new Error(`ช่วงเวลาที่เลือกทับซ้อนกับการจองอื่นสำหรับ '${resource}'`);
          }
      }
    }
  }
}

function findBookingRowById_(sheet, bookingId) {
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) throw new Error("Sheet is empty, cannot find booking.");
    
    const headers = data.shift();
    const bookingIdCol = headers.indexOf("BookingID");
    
    if (bookingIdCol === -1) throw new Error("BookingID column not found.");

    const rowIndexInMap = data.findIndex(r => String(r[bookingIdCol]) === String(bookingId)); 
    if (rowIndexInMap === -1) throw new Error(`ไม่พบข้อมูลการจอง ID: ${bookingId}`);
    
    return {
        row: data[rowIndexInMap],
        rowIndex: rowIndexInMap + 2, 
        headers: headers
    };
}

function getUserEmail_() {
    const user = Session.getActiveUser();
    const email = user ? user.getEmail() : '';
    
    if (!email) {
        throw new Error("ไม่สามารถระบุตัวตนผู้ใช้ได้ กรุณา Login Google Account");
    }
    
    return email;
}

function parseSheetDate_(sheetDate) {
    if (sheetDate instanceof Date && !isNaN(sheetDate)) {
        return sheetDate;
    }

    try {
        let date = new Date(sheetDate);
        if (!isNaN(date)) {
            if (date.getFullYear() > 2500) {
                date.setFullYear(date.getFullYear() - 543);
            }
            return date;
        }
    } catch (e) {}

    return null; 
}

function ensureHeader_(sheet, requiredHeaders) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const headersToAdd = [];
  
  requiredHeaders.forEach(reqHeader => {
    const currentHeadersTrimmed = currentHeaders.map(h => String(h).trim());
    if (!currentHeadersTrimmed.includes(reqHeader)) {
      headersToAdd.push(reqHeader);
    }
  });
  
  if (headersToAdd.length > 0) {
    const startColumn = currentHeaders.length + 1;
    sheet.getRange(1, startColumn, 1, headersToAdd.length).setValues([headersToAdd]);
  }
}

// ===============================================
// --- Admin Authentication (ส่วนตรวจสอบสิทธิ์) ---
// ===============================================

function checkIfUserIsAdmin_() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!userEmail) return false;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let adminSheet = ss.getSheetByName(SHEET_NAME_ADMINS);
    
    if (!adminSheet) {
      return false;
    }

    const lastRow = adminSheet.getLastRow();
    if (lastRow < 2) return false;
    
    const adminEmails = adminSheet.getRange(2, 1, lastRow - 1, 1).getValues()
                                .flat()
                                .map(email => String(email).toLowerCase().trim()) 
                                .filter(String);
                                
    return adminEmails.includes(userEmail);

  } catch (e) {
    return false;
  }
}

function checkIfUserIsAdminOrThrow_() {
  if (!checkIfUserIsAdmin_()) {
    throw new Error("Unauthorized: คุณไม่มีสิทธิ์ดำเนินการนี้ หรือคุณยังไม่มีสิทธิ์เข้าถึง Database");
  }
}

function getAdmins() {
  checkIfUserIsAdminOrThrow_(); 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const adminSheet = ss.getSheetByName(SHEET_NAME_ADMINS);
    if (!adminSheet || adminSheet.getLastRow() < 2) return [];
    
    return adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 1).getValues()
                       .flat()
                       .map(email => String(email).trim()) 
                       .filter(String);
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการดึงรายชื่อผู้ดูแล");
  }
}

function addAdmin(email) {
  checkIfUserIsAdminOrThrow_(); 
  if (!email || !email.includes('@')) {
    throw new Error("รูปแบบอีเมลไม่ถูกต้อง");
  }
  const trimmedEmail = email.trim().toLowerCase();
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const adminSheet = ss.getSheetByName(SHEET_NAME_ADMINS);
    const adminEmails = getAdmins().map(e => e.toLowerCase());

    if (adminEmails.includes(trimmedEmail)) {
      throw new Error(`อีเมล '${email}' เป็นผู้ดูแลอยู่แล้ว`);
    }

    adminSheet.appendRow([trimmedEmail]);
    return { success: true, message: "เพิ่มผู้ดูแลสำเร็จ", newAdmin: trimmedEmail };
  } catch (e) {
    throw new Error(e.message || "เกิดข้อผิดพลาดในการเพิ่มผู้ดูแล");
  }
}

function deleteAdmin(email) {
  checkIfUserIsAdminOrThrow_(); 
  const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  const emailToDelete = email.toLowerCase().trim();

  if (currentUserEmail === emailToDelete) {
    throw new Error("คุณไม่สามารถลบตัวเองออกจากรายชื่อผู้ดูแลได้");
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const adminSheet = ss.getSheetByName(SHEET_NAME_ADMINS);
    const data = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 1).getValues();
    
    const rowIndex = data.findIndex(row => String(row[0]).toLowerCase().trim() === emailToDelete);
    
    if (rowIndex === -1) {
      throw new Error(`ไม่พบผู้ดูแลด้วยอีเมล '${email}'`);
    }

    adminSheet.deleteRow(rowIndex + 2); 
    return { success: true, message: `ลบผู้ดูแล '${email}' สำเร็จ`, removedAdmin: email };
  } catch (e) {
    throw new Error(e.message || "เกิดข้อผิดพลาดในการลบผู้ดูแล");
  }
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
