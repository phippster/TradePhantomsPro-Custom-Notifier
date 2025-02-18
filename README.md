I wanted to be able to get notifications about trades that were relevant to me.  For instance I'm most interested in Swing trades and options.  By installing this script I will be able to get email notifications about updates to the trade picks sheet but ONLY for the trades I'm interested in.
This will reduce clutter in my email inbox and allow me to focus on what's most important to me.
If you're interested in having the same capability copy my script source code and install it.  Here's how.

1. Open a blank Google Spreadsheet
2. You can name it if you want.  As this will be your personal sheet that will send you customized notification. 

All of the following steps we will only do this very first time and never need to do them again.  So hang on but don't worry.
1. Click on Extensions and then select Apps Script
2. In the Editing window delete the existing blank function on line 1-4.
3. You might get a dialog window letting you know what email you are currently logged in with.  Not sure why but just click Okay.
4. Paste it my code.
5. Either click the Save button or press command-S on your keyboard.
6. Now you'll only need to do this one time, but it's important to make sure the function drop down shows "onOpen".  If it does, press the run button.
7. You will be prompted to authorize the script.  Click on the Review Permissions button.
8. Choose which account you are logged in with by clicking on it.
9. Since this script is just for us, it's not something I have submitted to Google for review.  So you want to click on the "Advanced" link in the lower left.
10. You want to continue so click the "Go to Untitle Project".  This will allow your script to run.
11. On the next screen click "Continue".  Basically Google is just letting you know what information the script can access.
12. On the next page click "Select All" to authorize your script.  
13. Scroll to the bottom and click Continue.
14. In rare cases you might see a warning at the bottom that says "This project requires access to your Google Account to run. Please try again and allow it this time."  This happens if you take too long to authorize everything we just reviewed.  If you see that just click the Run button again at the top to repeat.
15. Your script will run and you'll never need to do any of that again.


1. Now your sheet has been updated with the script and initialized.
2. You can close the Apps Script tab as you'll no longer need it.

Returning to your new sheet you'll see two tabs at the bottom.  The Filtered Trades tab is where you'll see the trades that meet your search criteria.  You don't want to modify the data in either of these tabs as it will be auto-generated whenever you run the script or by the automated background process.

To start you'll see all the same data as the Pick Sheet.    Now let's create some criteria to filter the results to just those trades we are interested in.  To do that go to the TradePhantamsPro Notifier menu and select Search Criteria Builder.

Here is where you can create your search criteria.  First pick a column from the file that you want to use in the filter.  Next enter the type of filter such as Equals (so the value must match exactly -- case insensitive) or contains and for numeric fields you can use the greater than or less than operators.  You can add an additional criteria by clicking the Add Criteria button.  If you have more than one criteria by default the search will make sure that all of the criteria match.  If you want to change that to either (like a logical OR) you can do that.  Click the Save Criteria button.  Be patient because when you change your search criteria the script will go grab the latest trades from the Pick sheet and then search through them to find any matching trades.  You'll see the data update on the Filtered Trades tab.

You may notice now that you also got an email showing which trades were added to your file.  From this point forward you'll get an email notification when any new trades are added that meet your criteria and if any of the existing trades are updated.

You can choose to run this script MANUALLY by going to the menu and selecting Run Notifier Once.

But you probably don't want to have to leave this sheet open on your computer and click Run Notifier Once all the time.  To solve for that I created a background process that will run the script on a regular time schedule.  The default is every 5 minutes but you can change the frequency between 1, 5, 10, 15, or 30 minutes.  To change the Frequency select Set Frequency from the Menu and enter your desire frequency.

After you've select the frequency you desire it's time to turn on the background process.  Do that by selecting Start Background Notifier from the menu.  It will now run in the background even after you close the sheet.

If you decide you want it to stop just open your sheet and select Stop Notifier from the menu.

As I mentioned you will get email notifications when things change.  Previously the best we could do was get a generic email from google stating that the TradePicks sheet had been updated.  Now you'll get a detailed email with the details on each new or updated trade.

That's all.  If you have questions please feel free to drop me a line.  I'm sure we'll run into some bugs here and there as I'm not perfect but I hope this helps.

Mark out!


