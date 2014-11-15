<?php
/**
* Excel 2003 XML PHP Parser

* PHP class for reading (parsing) a Microsoft Excel 2003 XML Spreadsheet and return it as a PHP array.

* The class use `simplexml` function.

* The returned array contains:


* ·         `arr_properties` as the excel page properties.

* ·         ` table_contents` an array of:

* ·         ` row_num` related row number

* ·         ` row_contents`

* o   ` row_num` related row number

* o   `col_num` related column number

* o   ` datatype` related column defined type

* o   ` value` the column content

* ·         ` count_contents` amount of row column with contents, `0` for empty row
 
*/
class excel2003XMLParser{
        
        protected static $instance; 
        protected $arr_excel;
        protected $arr_row;   
        
        public function __construct(){  
              $this->arr_excel  = array();
              $this->arr_row    = array();
        }
        
        /**
        * get instance of excel2003XMLParser
        * 
        */
         public static function GetInstance(){   
            if(empty(self::$instance)){
                self::$instance = new excel2003XMLParser();  
            }
            return self::$instance;
        }   
        /**
        * get attributes from xml element object
        * 
        * @param mixed $arr_attributes
        */
        private function get_attributes($arr_attributes){
            $arr_attribute_data = array();
            foreach( (array)$arr_attributes as $arr_attribute){
                $arr_attribute = (array)$arr_attribute;
                foreach($arr_attribute as $arr_attrib){
                    $attr_key = key( (array)$arr_attrib);
                    $arr_attribute_data[$attr_key] = $arr_attrib[$attr_key];
                }
            }
            return $arr_attribute_data;
        }   
        /**
        * parse XML file
        * 
        * @param mixed $url
        */
        public function loadXMLFile($url){
            $this->arr_excel = array(
                                    'arr_properties' => array(),
                                    'table_contents' => array()
                                );  
             if(!file_exists($url)){
                  return "Error - file not exist";
             }                        
            // assign simpleXML object
            if($simplexml_table = simplexml_load_file($url)){
                
                // valid XML 2003 spreadsheet                    
                $xmlns = $simplexml_table->getDocNamespaces();    
                if($xmlns['ss'] != 'urn:schemas-microsoft-com:office:spreadsheet'){
                    return "Error - file not valid XML 2003 spreadsheet";
                }
            } else {     
                // error loading file
                return "Error - reading file";
            }      
            // extract document properties
            $arr_properties = (array)$simplexml_table->DocumentProperties;
            $this->arr_excel['arr_properties'] = $arr_properties;
             // extract rows
            $rows       = $simplexml_table->Worksheet->Table->Row;
            $row_num    = 1;   
            // loop through all rows        
            foreach($rows as $row){   
                $cells              = $row->Cell;
                $row_attrs          = @$row->xpath('@ss:*');
                $arr_row_attrs      = $this->get_attributes($row_attrs);
                $this->arr_row      = array();
                $col_num            = 1;  
                $count_cell_data    = 0;
                // loop through all row's cells
                foreach($cells as $cell){  
                    // check whether ss:Index attribute exist
                    $cell_index = @$cell->xpath('@ss:Index'); 
                    $cell_index = (int)$cell_index[0]; 
                    // if exist, push empty value until the specified index
                    if($cell_index){
                        $loop_index     =  $cell_index - $col_num;
                        for($i = 0; $i < $loop_index; $i++){
                             // insert column data
                            $this->arr_row[$col_num] =  array(
                                                    'row_num' => $row_num,
                                                    'col_num' => $col_num,
                                                    'datatype' => '',
                                                    'value' => '',                            
                                                    'cell_attrs' => '',
                                                    'data_attrs' => ''
                                                );  
                            $col_num++;
                        }
                    }     
                    // get all cell and data attributes                
                    $cell_attrs     = @$cell->xpath('@ss:*');
                    $arr_cell_attrs = $this->get_attributes($cell_attrs);
                    
                    $data_attrs     = @$cell->Data->xpath('@ss:*');
                    $arr_data_attrs = $this->get_attributes($data_attrs);
                    $cell_datatype  = $arr_data_attrs['Type'];  
                    // extract data from cell
                    $cell_value     = trim( (string)$cell->Data );  
                    // insert column data
                    $this->arr_row[$col_num] =  array(
                                            'row_num'       => $row_num,                    
                                            'col_num'       => $col_num,
                                            'datatype'      => $cell_datatype,        
                                            'value'         => $cell_value,                    
                                            'cell_attrs'    => $arr_cell_attrs,
                                            'data_attrs'    => $arr_data_attrs
                                        );
                    if($cell_value){
                        $count_cell_data++;
                    } 
                    $col_num++;
               
                }  
                    // push row array
                $this->arr_excel['table_contents'][] = array(
                                                                'row_num'       => $row_num,
                                                                'row_contents'  => $this->arr_row,
                                                                'count_contents'=> $count_cell_data,
                                                                'row_attrs'     => $arr_row_attrs
                                                            );
                $row_num++;
            }                           
            // return array of excel data 
            return $this->arr_excel;
        }
    }
    

      
$obj_excel = excel2003XMLParser::GetInstance();
$arr_excel = $obj_excel->loadXMLFile('your-file.xml');
?>
 