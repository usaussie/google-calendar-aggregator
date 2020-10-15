<p align="center">

  <h3 align="center">Google Calendar Data Aggregator</h3>

  <p align="center">
    AppScript code to query your google calendar and pull data into a [Google Sheets](https://sheets.google.com), which can then be analyzed or visualized in [Google Data Studio](https://datastudio.google.com)
  </p>
</p>


## Why

I wanted to create a method for users to audit their own Google Calendar activity/attendees etc, without having to give access to a 3rd party application/service.

## Built With
This is a native AppScript solution...just a little appscript...then using the built-in triggers to run on a schedule to grab "yesterday's" data (can also run a one-time back-fill.

* [AppScript](https://script.google.com)

## Getting Started

* Copy the raw code from the Code.gs and email.html files to your appscript project
* Keep the file names the same
* Update the variables inside the Code.gs file to point to your results sheet, and tab name.
* Run the set_sheet_headers() function, which will prompt for permissions
* Accept the permissions (asking for access for your script to google drive and google calendar). 
* Run the get_yesterdays_events() function (once, or set a trigger)
* Look in your google sheet as the function is running, and you should see results being inserted 

<!-- CONTACT -->
## Contact

Nick Young - [@techupover](https://twitter.com/techupover)

Project Link: [https://github.com/usaussie/google-calendar-aggregator](https://github.com/usaussie/google-calendar-aggregator)
