// base config
var ss = SpreadsheetApp.getActiveSpreadsheet();
var ps = ss.getSheetByName("Team Planning");
var ts = ss.getSheetByName("Timeline");
var cs = ss.getSheetByName("settings");
var ts_datesStart, ts_baseDate, ts_baseDateVal;
var configDict = {};
var configSheetRngA1 = "B10:B22";
var domainColDict = { "App":"#99CC99", "Cloud":"#2BC2D3", 
                       "Cloud + App":"#FF0000", "Firmware":"#586791", 
                       "Test 1":"#586791", "Test 2":"#FF0000", 
                       "Test 3":"#FF0000" }

// get config stuff from the 'settings' sheet
function populateConfigDict()
{
  var configRng = cs.getRange("B10:B22");
  var configRngVals = configRng.getValues();
  for ( var i = 0; i < configRngVals.length; i++ )
  {
    configDict[i] = configRngVals[i][0];
  }
  ts_datesStart = configDict[1];
  ts_baseDate = ts.getRange(ts_datesStart);
  ts_baseDateVal = ts_baseDate.getValue();
}

// remove the resource alloc. weightages from timeline sheet to
// add updated ones from the roadmap sheet
function cleanUpResourceTimeline()
{
  var ts_resourceTimelineArea = ts.getRange(2, 3, 31, 520);
  ts_resourceTimelineArea.setValue(null);
}

// this function is responsible for updating the timeline sheet
// based on the start and end dates for each milestone.
function updateProjectTimeline()
{
  var b_baseRow = 32, b_baseCol = 2;
  var projectCodesRng = ps.getRange(configDict[2] + ":" + 
                                    configDict[2].replace(/[0-9]/g, '') + 
                                    String(configDict[11]));
  var projectCodes = projectCodesRng.getValues();
  var timelineCodeRng = ts.getRange(b_baseRow, b_baseCol, configDict[11]);
  var timelineCodeRngVals = timelineCodeRng.getValues();
  var datesRng = ts.getRange(configDict[0]);
  var datesRngVals = datesRng.getValues();
  var projectsDict = {};
  var domainRng = ps.getRange(configDict[3] + ":" + 
                              configDict[3].replace(/[0-9]/g, '') + 
                              String(configDict[11]));
  var domainVals = domainRng.getValues();
  var projDescriptRng = ps.getRange(configDict[4] + ":" + 
                                    configDict[4].replace(/[0-9]/g, '') + 
                                    String(configDict[11]));
  var projDescriptVals = projDescriptRng.getValues();
  // cleaning up the timeline sheet to create a fresh timeline
  var timelineSheetPaintArea = ts.getRange(b_baseRow, b_baseCol + 1, 120, 520);
  timelineSheetPaintArea.setBackground("#FFFFFF");
  timelineSheetPaintArea.setValue("");
  // looping through project codes and creating timelines for each
  for ( var i = 0; i < projectCodesRng.getNumRows(); i++ )
  {
    projectsDict[i] = domainColDict[domainVals[i]];
  }
  for ( var j = 0; j < projectCodesRng.getNumRows(); j++ )
  {
    for ( var i = 0; i < projectCodesRng.getNumRows(); i++ )
    {
      var projCodeVal = projectCodes[j][0];
      var timelineCodeVal = timelineCodeRngVals[i][0];
      var timelineCellMod = ts.getRange(b_baseRow + i, b_baseCol + 1);
      if ( timelineCodeVal == projCodeVal )
      {
        var startDateCell = projectCodesRng.getCell(j + 1, 1).offset(0, configDict[9]);
        var startDateVal = startDateCell.getValue();
        if ( startDateVal != "" )
        {
          var endDateCell = projectCodesRng.getCell(j + 1, 1).offset(0, configDict[10]);
          var endDateVal = endDateCell.getValue();
          if ( endDateVal != "" )
          {
            var timelineSheetStartDateLoc = datesRng.getCell(1, 1).getColumn();
            var startCol = 0, endCol = 0;
            for ( var k = 0; k < datesRng.getNumColumns(); k++ )
            {
              var val = datesRngVals[0][k];
              if ( val.getTime() == startDateVal.getTime() )
              {
                startCol = k;
              }
              if ( val.getTime() == endDateVal.getTime() )
              {
                endCol = k;
                break;
              }
            }
            var timelineStartCell = timelineCellMod.offset(0, startCol);
            timelineStartCell.setValue(projDescriptVals[j]);
            var timelineEndCell = timelineCellMod.offset(0, endCol);
            var startCellA1Notation = timelineStartCell.getA1Notation();
            var endCellA1Notation = timelineEndCell.getA1Notation();
            var timelineRng = ts.getRange(startCellA1Notation + ":" 
                                          + endCellA1Notation);
            timelineRng.setBackgroundColor(projectsDict[j]);
            break;
          }
        }
      }
    }
  }
}

function searchStringInRange(rng, txt)
{
  var txtFound = false;
  var foundCell;
  var vals = rng.getValues();
  var count = 0;
  do
  {
    if ( vals[0][count] == txt )
    {
      foundCell = rng.getCell(1, count + 1);
      txtFound = true;
    }
    count += 1;
  }
  while ( txtFound == false );
  if ( txtFound == true )
  {
    return foundCell;
  }
  else
  {
    return null;
  }
}

// this function is responsible for plotting the workload timeline for 
// each team member in the project (need to split this into smaller chunks)
function updateResourcesTimeline()
{
  cleanUpResourceTimeline();
  // below range for team members would need to be increased to
  // accommodate more resources.
  var projectCodesRng = ps.getRange(configDict[2] + ":" + 
                                    configDict[2].replace(/[0-9]/g, '') + 
                                    String(configDict[11]));
  var projectCodes = projectCodesRng.getValues();
  var test = configDict[2] + ":" + configDict[2].replace(/[0-9]/g, '') + 
                                    String(configDict[11])
  var timelineTeamRng = ts.getRange(configDict[8] + ":" + 
                                    configDict[8].replace(/[0-9]/g, '') + 
                                    String(2 + configDict[12]));
  var timelineTeamRngVals = timelineTeamRng.getValues();
  var timelineTeamBaseCell = timelineTeamRng.getCell(1, 1);
  var planningTeamFirstCell = ps.getRange(configDict[5]);
  var planningTeamLastCell = planningTeamFirstCell.offset(0, configDict[12] - 1);
  var planningTeamLastCellA1Notation = planningTeamLastCell.getA1Notation();
  var planningTeamRng = ps.getRange(configDict[5] + ":" + planningTeamLastCellA1Notation);
  var timelineTeamVals = timelineTeamRng.getValues();
  var planningTeamVals = planningTeamRng.getValues();
  var startDateRng = ps.getRange(configDict[6] + ":" + 
                                 configDict[6].replace(/[0-9]/g, '') + 
                                 String(configDict[11]));
  var startDateVals = startDateRng.getValues();
  var endDateRng = ps.getRange(configDict[7] + ":" + 
                               configDict[7].replace(/[0-9]/g, '') + 
                               String(configDict[11]));
  var endDateVals = endDateRng.getValues();
  var timelineDatesHeaderRng = ts.getRange(configDict[0]);
  var timelineDatesHeaderVals = timelineDatesHeaderRng.getValues();
  for ( var i = 0; i < planningTeamRng.getNumColumns(); i++ )
  {
    var allocDict = {}; // the main conatiner for resource allocation for each milestone
    // below array with contain the aggregated weightages for a day 
    // across multiple milestones for the current resouce. This is
    // required to push the array information all at once into the
    // determined range for the users timeline allocation on the
    // timeline sheet.
    var timelineAllocArr = [];
    resourceCell = planningTeamRng.getCell(1, i + 1);
    resourceName = resourceCell.getValue();
    // iterate through the allocations for the current resource and
    // collect the start and end dates for each assignment along with
    // the weightage for that particular assignment,
    var allocStartCell = resourceCell.offset(1, 0);
    var allocStartCellA1Notation = allocStartCell.getA1Notation();
    var columnNameText = allocStartCellA1Notation;
    columnNameText = columnNameText.replace(/[0-9]/g, '');
    var allocEndCellA1Notation = columnNameText + String(configDict[11]);
    var allocRng = ps.getRange(allocStartCellA1Notation + ":" + allocEndCellA1Notation);
    var allocVals = allocRng.getValues();
    //var destinationAllocRng = timelineSheet.getRange();
    for ( var j = 0; j < allocRng.getNumRows(); j++ )
    {
      if ( allocVals[j] != "" )
      {
        // get the project code, the allocation weightage, and the
        // allocation timeframe for the resource and push it into
        // a dictionary.
        if (( startDateVals[j] != null ) && (endDateVals[j] != null))
        {
          allocDict[projectCodes[j]] = [startDateVals[j], endDateVals[j], allocVals[j]];
        }
      }
    }
    // now to plot allocation values for the team member on the
    // timeline sheet.
    var timelineMemberRowIndex = 0;
    for ( var k = 0; k < timelineTeamRng.getNumColumns(); k++ )
    {
      if ( resourceName == timelineTeamVals[k] )
      {
        timelineMemberRowIndex = k;
        break;
      }
    }
    var timelineCurrMemberCell = timelineTeamBaseCell.offset(timelineMemberRowIndex, 0);
    var dateArr = getDateLimitsArrFromDict(allocDict);
    var testDate1 = dateArr[0];
    var testDate2 = dateArr[dateArr.length - 1];
    var numberOfDays = getNumberOfDaysBetweenDates( testDate1, testDate2 );
    // now we need to iterate over the project allocations for the resource,
    // get the start and end date for each along with the weightage, and then
    // add the weigtages at the correct positions along dateArr array.
    if ( numberOfDays != 0 )
    {
      for ( var l = 0; l <= numberOfDays; l++ )
      {
        timelineAllocArr.push([0]);
      }
      for ( var m = 0; m < Object.keys(allocDict).length; m++ )
      {
        var projCode = Object.keys(allocDict)[m];
        var projStart = allocDict[projCode][0][0];
        var projEnd = allocDict[projCode][1][0];
        var projWeight = allocDict[projCode][2][0];
        var numDaysBaseAndProjStart = getNumberOfDaysBetweenDates(dateArr[0], projStart);
        var numDaysBaseAndProjEnd = getNumberOfDaysBetweenDates(dateArr[0], projEnd);
        for ( var n = 0; n <= (numDaysBaseAndProjEnd - numDaysBaseAndProjStart); n++ )
        {
          timelineAllocArr[numDaysBaseAndProjStart + n][0] = 
            timelineAllocArr[numDaysBaseAndProjStart + n][0] + projWeight;
        }
      }
      // now to 'paint' the resource's timeline...
      var cellRngA1 = getResourceTimelineStartAndEnd(ts_baseDateVal, dateArr[0], 
                                                     dateArr[dateArr.length - 1], 
                                                     resourceName, timelineTeamRngVals);
      var cellRng = ts.getRange(cellRngA1);
      var cellRngVals = cellRng.getValues();
      var arrPack = [timelineAllocArr];
      cellRng.setValues(arrPack);
    }
  }
}

// this function returns the A1 notation for the start and end cells
// on the timeline for resource allocation on the timelinesheet, given
// that the overall start and end dates are provided along with resource
// name
function getResourceTimelineStartAndEnd(baseDate, startDate, endDate, 
                                        name, timelineTeamRngVals)
{
  var offsetStart = getNumberOfDaysBetweenDates( baseDate, startDate );
  var offsetEnd = getNumberOfDaysBetweenDates( baseDate, endDate );
  var resourceIndex = 0;
  for ( var i = 0; i < timelineTeamRngVals.length; i++ )
  {
    if ( name == timelineTeamRngVals[i] )
    {
      resourceIndex = i;
      break;
    }
  }
  // assuming that the names start on the second row on the timline sheet
  var resourceRowNum = resourceIndex + 2;
  var strtCell = ts_baseDate.offset(1, offsetStart);
  var strtCellA1 = strtCell.getA1Notation();
  var strtCellTrim = strtCellA1.replace(/[0-9]/g, '');
  var endCell = ts_baseDate.offset(1, offsetEnd);
  var endCellA1 = endCell.getA1Notation();
  var endCellTrim = endCellA1.replace(/[0-9]/g, '');
  var concatStr = strtCellTrim + String(resourceRowNum) + ":" + endCellTrim + 
    String(resourceRowNum);
  return concatStr;
}

// function assumes date1 < date2
function getNumberOfDaysBetweenDates(date1, date2)
{
  var datesEqual = false;
  var days = 0;
  var local_date1 = new Date(date1);
  var local_date2 = new Date(date2);
  if ( ( date1 ) && ( date2 ) )
  {
    if (local_date1.getTime() != local_date2.getTime())
    {
      do
      {
        local_date1.setDate(local_date1.getDate()+1);
        days += 1;
        if ( local_date1.getTime() == local_date2.getTime() )
        {
          datesEqual = true;
        }
      }
      while ( datesEqual == false );
    }
  }
  return days;
}

// this will take the allocation dictionary and extract the minimum 
// and maximum dates for the entire allocation for the project resource, 
// and return an array of length equal to the number of dates the 
// resource would be allocated. This array would then be used to determine
// where exactly additional allocations are made, and add allocation
// weightages to the respective indices of the array. The array will 
// then be used to push the timeline for the resource on the chart.
function getDateLimitsArrFromDict(allocDict)
{
  var dateArr = [];
  for ( var m = 0; m < Object.keys(allocDict).length; m++ )
  {
    var startDate = allocDict[Number(Object.keys(allocDict)[m])][0][0];
    var endDate = allocDict[Number(Object.keys(allocDict)[m])][1][0];
    if ( startDate )
    {
      dateArr.push( startDate );
    }
    if ( endDate )
    {
      dateArr.push( endDate );
    }
  }
  // now to sort the dates array...
  var swapped = false;
  var indexOfLastUnsorted = 0;
  do
  {
    swapped = false;
    for ( var k = 0; k < dateArr.length; k++ )
    {
      var item1 = dateArr[k];
      var item2 = dateArr[k+1];
      if ( item1 > item2 )
      {
        var temp = dateArr[k];
        dateArr[k] = dateArr[k+1];
        dateArr[k+1] = temp;
        swapped = true;
      }
    }
  }
  while ( swapped == true );
  var timelineStartDate = dateArr[4];
  var timelineEndDate = dateArr[dateArr.length - 1];
  return dateArr;
}


// this function is responsible for creating two-way links between the
// project codes on the main roadmap sheet, and the project codes on
// the timeline sheet. this would allow easier navigation by simply
// clicking the link inside the cell to jump to the corresponding
// cell in the other sheet.
function establishTwoWayLinks()
{
  
}

// this function will help to remove spaces between milestones on the
// timeline sheet in case there are milestones which do not have dates
// defined and occupy space
function reorderTimelineSheetCodes()
{
  
}

// this function will order milestones on the timeline sheet based on
// common themes and association to give a better picture of roadmap
// for different themes
function reorderByTheme()
{
  
}

// this function does everything requierd to update the timeline sheet
function main()
{
  populateConfigDict();
  updateResourcesTimeline();
  updateProjectTimeline();
}