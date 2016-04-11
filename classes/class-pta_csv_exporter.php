<?php
/**
* Exports reports from PTA plugins in CSV format
* This class is shared/used by several PTA plugins
*/
if ( ! defined( 'ABSPATH' ) ) exit; // Exit if accessed directly

class PTA_SUS_CSV_EXPORTER {

    public function  __construct() {

        // For exporting data from various admin pages as Excel CSV
        add_action('admin_init', array($this, 'export_csv'));

    }


    /*****************************************************************************************
     ********************************** EXPORT CSV FUNCTIONS ********************************* 
     *****************************************************************************************/

    /**
     * Exports previously prepared report arrays as a CSV Excel file
     * This little function is called by the admin_init wordpress hook
     * It checks if an export link was clicked on, validates the Nonce, and then 
     * echos the export_program function output to create the CSV file
     * Needs to be hooked to admin_init so that it can write file headers before
     * Wordpress outputs its own header info
     * 
     * @return CSV exports a csv file to the broswer for open/save
     */
    public function export_csv() {
        $export = isset($_REQUEST['pta-action']) ? $_REQUEST['pta-action'] : '';
        if ($export == 'export') {
            check_admin_referer('pta-export');
            if (!current_user_can('manage_options') && !current_user_can('manage_pta'))  {
                wp_die( __( 'You do not have sufficient permissions to access this page.' ) );
            }
            
            // Grab the option name where the report is stored, or use the default location
            $data = isset($_REQUEST['data']) ? $_REQUEST['data'] : 'pta_reports';

            if ( !$report = get_option($data) ) {
                wp_die( __( 'Invalid Report Data!' ) );
            }
            echo $this->export_report($report);
            exit;
        }
        
    }

    /**
     * This function creates the header info and CSV output for the above function
     * @param  wordpress option 'pta_reports' Report data stored in the Wordpress Option table
     * @return headers             header info so browser knows we're sending it a CSV file
     * @return string   the CSV data to go with the file
     */
    public function export_report($report) {

        $csv = '';

        header("Content-type: text/csv");
        header("Content-Disposition: attachment; filename=report-".date('Ymd-His').".csv");
        header("Pragma: no-cache");
        header("Expires: 0");

        $headers = array(
            "Presenter",
            "Date",
            "Start Time",
            "End Time",
            "Participant",
            "Email",
            "Phone"
        );

        $count = 1;
        foreach ($report as $data) { 

            if ( 1 == $count ) { // Outputs the column headers from the array keys

                $key_array = array_keys($data);
                $num_keys = count($key_array);
                
                $i = 1;
                foreach ( $headers as $header ) {
                    $csv .= '"' . $header . '"';
                    if ($i < $num_keys) {
                        $csv .= ',';
                    }
                    $i++;
                }
                $csv .= "\n";
            }

            // If name is empty, do not output to csv file
            if( $data['Volunteer Name'] === 'empty' && $count > 2 ) {
                break;
            }
            
            $i = 1;
            $task = $data['Task/Item'];
            $date = $data['Task Date']; // Grab the date
            $start_time = $data['Start Time'];
            $end_time = $data['End Time'];
            foreach ( $data as $key => $value ) {
                if( ($value === $task && $count > 1) ||
                    ($value === $date && $count > 1) || 
                    ($value === $start_time && $count > 1) ||
                    ($value === $end_time && $count > 1) ) { // Do not repeat duplicate data
                    $value = "";
                }
                if( $key == "Item Qty" || $key == "Item Details" ) { // Do not export these 2 items
                    break;
                }
                $csv .= '"' . $this->clean_csv($value) . '"';
                    if ($i < $num_keys) {
                        $csv .= ',';
                    }
                    $i++;
            }
            $csv .= "\n";
            $count++;
        }
        return $csv;
    }   

    /**
     * Small helper function to get any quotes in proper format
     * @param  string $value input string to clean
     * @return string        cleaned value
     */
    private function clean_csv($value)
    {
        // let's clean any html breaks out as well
        $value = str_replace('<br/>', ', ', $value);
        $value = str_replace('<br />', ', ', $value);
        // let's also convert any underscores to spaces
        $value = str_replace('_', ' ', $value);
        $value = str_replace('"', '""', $value);
        // Strip any remaining html tags
        $value = strip_tags($value);
        return $value;
    }

} // End Class

$pta_sus_csv_exporter = new PTA_SUS_CSV_EXPORTER();

/* EOF */
