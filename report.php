<?php

/**
 * Plugin Name: CS Activities Report
 * Plugin URI:  https://wordpress.org/plugins/#/
 * Author:      Lazutina
 * Author URI:  https://profiles.wordpress.org/#
 * Version:     1.0
 * Text Domain: #
 * Domain Path: #
 */

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
require_once __DIR__ . '/PhpSpreadsheet/src/Bootstrap.php';
// Exit if accessed directly
defined( 'ABSPATH' ) || exit;

class CS_Activity_Report{
    static function init(){
        //add_action('admin_menu', array(__CLASS__,'add_menu_page'), 100);
        add_action('init',array(__CLASS__,'initial'), 10);
        add_filter('user_has_cap',array(__CLASS__,'user_has_cap'), 10, 1);
        add_action('admin_menu',array(__CLASS__,'admin_menu'), 100);
        add_action('wp',array(__CLASS__,'generate'), 10);
    }
    static function initial(){
        self::add_teacher_role();
    }
    static function generate(){
        if(isset($_GET['export_action'])&&$_GET['export_action'] == 'Export'){
            self::excel_generate();
            die;
        }
    }
    private static function get_students(){
        $user_ids = array();
		if(function_exists('search_students_by')){
			$first_name = null;
			$last_name = null;
			$year_id = 0;
			$home_id = 0;
			if($_GET['first_name'])$first_name = $_GET['first_name'];
			if($_GET['last_name'])$last_name = $_GET['last_name'];
			if($_GET['_year'])$year_id = $_GET['_year'];
			if($_GET['_house'])$home_id = $_GET['_house'];
			$user_ids = search_students_by($first_name,$last_name,$year_id,$home_id);
		}
        return $user_ids;
    }
    private static function excel_generate()
    {
        global $wp_query;

        $activities = $wp_query->posts;
        $posts = array();
        foreach($activities as $activity){
            if($activity->post_parent>0){
                if(isset($posts[$activity->post_parent])){
                    $posts[$activity->post_parent][] = $activity;
                }else{
                    $posts[$activity->post_parent] = array($activity);
                }
            }else{
                $posts[$activity->ID] = array($activity);
            }
        }
        $user_ids = self::get_students();
        // Create new Spreadsheet object
        $spreadsheet = new Spreadsheet();
        
        // Set document properties
        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
            ->setLastModifiedBy('Maarten Balliauw')
            ->setTitle('Office 2007 XLSX Test Document')
            ->setSubject('Office 2007 XLSX Test Document')
            ->setDescription('Test document for Office 2007 XLSX, generated using PHP classes.')
            ->setKeywords('office 2007 openxml php')
            ->setCategory('Test result file');
        
        // Add some data
        $spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Student First Name')
            ->setCellValue('B1', 'Student Last Name')
            ->setCellValue('C1', 'Year of graduation')
            ->setCellValue('D1', 'House');
        $index = 0;    
        foreach($posts as $post_id=>$activities){
            $spreadsheet->setActiveSheetIndex(0)->setCellValue(chr(69+$index).'1', $activities[0]->post_title);
            $index++;
        }
        foreach($user_ids as $index=>$user_id){
            $first_name = get_user_meta($user_id,'first_name',true);
            $last_name = get_user_meta($user_id,'last_name',true);
            $years = wp_get_terms_for_user($user_id, 'user-group');
            $year = get_year_graduate_user($user_id);
            $houses = wp_get_terms_for_user($user_id, 'user-type');
            if(isset($houses[0]->name))$house = $houses[0]->name;
            else $house = "";
            $spreadsheet->setActiveSheetIndex(0)
                ->setCellValue('A'.($index+2), $first_name)
                ->setCellValue('B'.($index+2), $last_name)
                ->setCellValue('C'.($index+2), $year)
                ->setCellValue('D'.($index+2), $house);
            $post_index=0;
            foreach($posts as $post_id=>$activities){
                foreach($activities as $activity){
                    if($activity->post_author == $user_id){
                        $approved_date = get_post_meta($activity->ID,'_activity_date',true);
                        if($approved_date){
                            $spreadsheet->setActiveSheetIndex(0)->setCellValue(chr(69+$post_index).($index+2), $approved_date);
                        }
                    }
                }
                $post_index++;
            }
        }
        // Rename worksheet
        $spreadsheet->getActiveSheet()->setTitle('Activity');
        
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);
        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="Activity_Export.xlsx"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        
        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0
        
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    }
    private static function force_download($filename) {
        $filedata = @file_get_contents($filename);
    
        // SUCCESS
        if ($filedata)
        {
            // GET A NAME FOR THE FILE
            $basename = basename($filename);
    
            // THESE HEADERS ARE USED ON ALL BROWSERS
            header("Content-Type: application-x/force-download");
            header("Content-Disposition: attachment; filename=$basename");
            header("Content-length: " . (string)(strlen($filedata)));
            header("Expires: ".gmdate("D, d M Y H:i:s", mktime(date("H")+2, date("i"), date("s"), date("m"), date("d"), date("Y")))." GMT");
            header("Last-Modified: ".gmdate("D, d M Y H:i:s")." GMT");
    
            // THIS HEADER MUST BE OMITTED FOR IE 6+
            if (FALSE === strpos($_SERVER["HTTP_USER_AGENT"], 'MSIE '))
            {
                header("Cache-Control: no-cache, must-revalidate");
            }
    
            // THIS IS THE LAST HEADER
            header("Pragma: no-cache");
    
            // FLUSH THE HEADERS TO THE BROWSER
            flush();
    
            // CAPTURE THE FILE IN THE OUTPUT BUFFERS - WILL BE FLUSHED AT SCRIPT END
            ob_start();
            echo $filedata;
        }
    
        // FAILURE
        else
        {
            die("ERROR: UNABLE TO OPEN $filename");
        }
    }
    static function add_teacher_role(){
        $roles = get_editable_roles();
        $exist = true;
        foreach($roles as $key=>$rule){
            if($key == 'teacher'){
                $exist = false;
            }
        }
        if($exist || true){
            add_role(
                'teacher',
                __( 'Teacher' ),
                array(
                    'read'         => true,  // true allows this capability
                    'edit_other_activities'   => true,
                    'edit_activities'   => true,
                )
            );  
            foreach ($GLOBALS['wp_roles']->role_objects as $key => $role) {
                if (isset($roles[$key]) && $role->has_cap('edit_posts')) {
                    $role->add_cap('read_activity');
                }
            }            
        }
    }
    static function admin_menu(){
        global $menu;
        $user = wp_get_current_user();
        if ( in_array( 'teacher', (array) $user->roles ) ) {
            foreach($menu as $index=>$item){
                if($item[1]=='edit_posts' || $item[1]=='wpcf7_read_contact_forms'){
                    unset($menu[$index]);
                }
            }
        }
    }
    static function add_menu_page(){
        add_menu_page('Activity Report','Activity Report','edit_posts','cs-activity-report',array( __CLASS__, 'report' ),null,76);
    }
    static function user_has_cap($capabilities){
        if(isset($_GET['post_type'])&&$_GET['post_type']=='cs-activity')$capabilities['edit_posts'] = true;
        return $capabilities;
    }
    static function report(){
        
    }
}
CS_Activity_Report::init();