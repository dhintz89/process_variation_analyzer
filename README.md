# process_variation_analyzer
VBA Macro Tool using Visio and Excel to analyze and report on process variation
*Required to be used in conjunction with a specific Visio Variance Template (which includes Shape Report) and Stencil.*

## Instructions:

1. After Process Doc is completed, partner with peer PA for variation analysis
2. Analyzing PA compares to SOP for same process
	* Determine meaningful vs. arbitrary variation
		* *Arbitrary: “Distribute Bookings File” vs. “Send Bookings Report”*
		* *Meaningful: “Receive Bookings File” vs. “Search for Recent Bookings in Client System”*
	* Highlight Meaningful Variation in Map (Use CTRL + H to highlight):
		* *Specs (for reference only): Line Weight = “1.5 pt” (exact); Line Color = Light Blue (not exact)*
		* **Note: can only highlight one box at a time, do not multi-select**
	* Do nothing for Arbitrary Variation
3. Once Complete, Dbl-Click “Variance Check” Button on Instructions Page
	* **Important: Before running, ensure there are no Excel Processes running on your PC – open Task Manager 	(Ctrl+Shift+Esc) and scroll through Apps and Background Processes ending any rogue Excel processes**
4. Macro Runs (don’t touch computer while macro runs):
	* Formats Visio shape data for download
	* Exports formatted data into Excel Variance Analysis Template Data Page
	* Calculates and formats Excel data into presentation view
	* Notifies user of end of process
5. Discuss results with Mapping PA and PA Lead
