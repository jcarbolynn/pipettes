# pipettes

using google apps script/google sheets
tracks when pipettes need to be calibrated or verified (using a scale to confirm that they are depositing correct amounts)
if the date for calibration or verification is within a week of whn the script is run (each monday), it will send an email to the person using the pipette of the date of verification and calibration and which pipette needs the maintenance
also sends the same email to the supervisor and me
