<?php header( 'Content-Type: text/html; charset=utf-8' );
require_once( 'PHPExcel.php' );
// Подключаем класс для вывода данных в формате excel
require_once( 'PHPExcel/Writer/Excel5.php' );
include 'PHPExcel/IOFactory.php';


$input_folder = 'input';
$db_connect   = array(
	'db_name'  => 'call_center_db',
	'login'    => 'root',
	'password' => 'root',
	'host'     => 'localhost',
);


class CallCenterParser {


	protected $db_name;
	protected $login;
	protected $password;
	protected $host;
	public    $link;

	/**
	 * CallCenterDB constructor.
	 * @param $db_connect
	 */
	public function __construct( $db_connect ) {
		$this->db_name  = $db_connect['db_name'];
		$this->login    = $db_connect['login'];
		$this->password = $db_connect['password'];
		$this->host     = $db_connect['host'];
		$this->link     = new mysqli( $this->host, $this->login, $this->password, $this->db_name );
		if ( $this->link->connect_error ) {
			die( 'Connect Error (' . $this->link->connect_errno . ') ' . $this->link->connect_error );
		}
	}

	/**
	 * Добавление необходимых таблиц для проекта
	 * @param int $debug
	 */
	public function insertTables( $debug = 0 ) {
		$sql = "CREATE PROCEDURE insertTables() 
		    BEGIN ";

		//Таблица адресов

		$sql .= "CREATE TABLE IF NOT EXISTS wp_address (
		  id int(11) NOT NULL AUTO_INCREMENT,
		  address varchar(510) NOT NULL,
		  PRIMARY KEY (id)
		) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1; ";

		// Таблица звонков

		$sql .= "CREATE TABLE IF NOT EXISTS wp_calldata (
				  id int(11) NOT NULL AUTO_INCREMENT,
				  call_id int(11) NOT NULL,
				  call_address varchar(255) NOT NULL,
				  call_sector int(11) NOT NULL,
				  call_pass varchar(255) NOT NULL,
				  call_porch varchar(255) NOT NULL,
				  call_floor int(11) NOT NULL,
				  doc_id int(11) NOT NULL,
				  doc_FIO varchar(255) NOT NULL,
				  call_comment text NOT NULL,
				  call_operator varchar(255) NOT NULL,
				  call_operator_id int(11) NOT NULL,
				  PRIMARY KEY (id)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1; ";

		// Таблица докторов

		$sql .= "CREATE TABLE IF NOT EXISTS `wp_doctors` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `doctor_FIO` varchar(255) NOT NULL,
				  `doctor_spec` varchar(255) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// Таблица пациентов

		$sql .= "CREATE TABLE IF NOT EXISTS `wp_patients` (
				  `id` bigint(20) NOT NULL AUTO_INCREMENT,
				  `FIO` varchar(255) NOT NULL,
				  `Birthday` date NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// Таблица адресов

		$sql .= "CREATE TABLE IF NOT EXISTS `wp_patient_address` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `address_id` int(11) NOT NULL,
				  `patient_id` int(11) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// Таблица данных пациентов
		$sql .= "CREATE TABLE IF NOT EXISTS `wp_patient_data` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `patient_id` int(11) NOT NULL,
				  `data` varchar(510) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// Таблица телефонов

		$sql .= "CREATE TABLE IF NOT EXISTS `wp_patient_phones` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `patient_phone` varchar(255) NOT NULL,
				  `patient_id` int(11) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// Таблица полисов
		$sql .= "CREATE TABLE IF NOT EXISTS `wp_patient_policy` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `patient_id` int(11) NOT NULL,
				  `policy` varchar(255) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// Таблица филиалов по адресам

		$sql .= "CREATE TABLE IF NOT EXISTS `wp_sector_address` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `sector` int(11) NOT NULL,
				  `address_id` int(11) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";

		// таблица докторов и филиалов

		$sql .= "CREATE TABLE IF NOT EXISTS `wp_sector_doctor` (
				  `id` int(11) NOT NULL AUTO_INCREMENT,
				  `sector` int(11) NOT NULL,
				  `doctor_id` int(11) NOT NULL,
				  PRIMARY KEY (`id`)
				) ENGINE=InnoDB DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;";


		$sql .= "END; ";

		if ( !$this->link->query( "DROP PROCEDURE IF EXISTS insertTables" ) ||
		     !$this->link->query( $sql ) ) {
			echo "Не удалось создать хранимую процедуру: (" . !$this->link->errno . ") " . !$this->link->error;
		}
		if ( !$this->link->query( "CALL insertTables()" ) ) {
			echo "Не удалось вызвать хранимую процедуру: (" . !$this->link->errno . ") " . !$this->link->error;
		}
		if ( $debug == 1 ) {

			echo $sql . "<br>";
			echo "<pre>";
			var_dump( $this->link );
			echo "</pre>";


		}

	}

	/**
	 * Очистка таблиц
	 */
	public function clearTables( $debug = 0 ) {
		$sql = "CREATE PROCEDURE clearTables() 
		    BEGIN ";
		$sql .= "DELETE FROM wp_address; ";
		$sql .= "DELETE FROM wp_calldata; ";
		$sql .= "DELETE FROM wp_doctors; ";
		$sql .= "DELETE FROM wp_patients; ";
		$sql .= "DELETE FROM wp_patient_address; ";
		$sql .= "DELETE FROM wp_patient_data; ";
		$sql .= "DELETE FROM wp_patient_phones; ";
		$sql .= "DELETE FROM wp_patient_policy; ";
		$sql .= "DELETE FROM wp_sector_address; ";
		$sql .= "DELETE FROM wp_sector_doctor; ";
		$sql .= "END; ";


		if ( !$this->link->query( "DROP PROCEDURE IF EXISTS clearTables" ) ||
		     !$this->link->query( $sql ) ) {
			echo "Не удалось создать хранимую процедуру: (" . !$this->link->errno . ") " . !$this->link->error;
		}
		if ( !$this->link->query( "CALL clearTables()" ) ) {
			echo "Не удалось вызвать хранимую процедуру: (" . !$this->link->errno . ") " . !$this->link->error;
		}
		if ( $debug == 1 ) {

			echo $sql . "<br>";
			echo "<pre>";
			var_dump( $this->link );
			echo "</pre>";


		}

	}

	/**
	 * Функция экспорта в массив данных из excel файла
	 */
	public function getCitizens() {
		global $input_folder;

		$citizens   = array();
		$req_fields = array(
			'fio'      => array( 'letter' => 'B', 'function' => 'escapeString' ),
			'birthday' => array( 'letter' => 'E', 'function' => 'prepareDate' ),
			'policy'   => array( 'letter' => 'C', 'function' => 'preparePolicy' ),
			'sector'   => array( 'letter' => 'G', 'function' => 'convertInt' ),
			'address'  => array( 'letter' => 'I', 'function' => 'escapeString' ),
			'phone'    => array( 'letter' => 'J', 'function' => 'preparePhone' ),
		);
		foreach ( glob( $input_folder . '/*.xls*' ) as $inputFileName ) {
			$pattern = '/' . $input_folder . '\/[^~$].*/';

			preg_match( $pattern, $inputFileName, $matches );
			if ( !$matches ) {
				continue;
			}

//  Read your Excel workbook
			try {
				$inputFileType = PHPExcel_IOFactory::identify( $inputFileName );
				$objReader     = PHPExcel_IOFactory::createReader( $inputFileType );
				$objPHPExcel   = $objReader->load( $inputFileName );

			} catch ( Exception $e ) {
				die( 'Error loading file "' . pathinfo( $inputFileName, PATHINFO_BASENAME ) . '": ' . $e->getMessage() );
			}
			$sheet         = $objPHPExcel->getSheet( 0 );
			$highestRow    = $sheet->getHighestRow();
			$highestColumn = $sheet->getHighestColumn();
			$rowData       = array();
			for ( $i = 6; $i < $highestRow; $i++ ) {

				foreach ( $req_fields as $field => $data ) {
					$cellValue = strval( $sheet->getCell( $data['letter'] . $i )->getValue() );
					$data['function'] != 'none' ? $cellValue = call_user_func_array( array( $this, $data['function'] ), array( $cellValue ) ) : $cellValue;
					$rowData[$i - 6][$field] = $cellValue;
				}
				if ( $i == 1005 ) {
					//break;
				}

			}
			echo "<pre>";
			print_r( $rowData );
			echo "</pre>";
//Избавляемся от внешнего массива


		}
		return $rowData;
	}

	/**
	 * Функция вставки данных в БД
	 *
	 */
	public function insertData( $data_arr, $debug = 0 ) {
		$query = "CREATE PROCEDURE insertRelations( 
					IN addressID INT(11),
					IN patientID INT(11),
                    IN phoneNum VARCHAR(255),
                    IN patientData VARCHAR(255),
                    IN userPolicy VARCHAR(255)
                    ) 
		    BEGIN 
		         INSERT INTO wp_patient_address (address_id,patient_id) VALUES (addressID,patientID); 
		         INSERT INTO wp_patient_policy (policy,patient_id) VALUES (userPolicy,patientID); 
		         INSERT INTO wp_patient_phones (patient_phone,patient_id) VALUES (phoneNum,patientID); 
		         INSERT INTO wp_patient_data (data,patient_id) VALUES (patientData,patientID); 
		    
		    END;
		    
		    ";
		if ( !$this->link->query( "DROP PROCEDURE IF EXISTS insertRelations" ) ||
		     !$this->link->query( $query ) ) {
			echo "Не удалось создать хранимую процедуру: (" . !$this->link->errno . ") " . !$this->link->error;
		}


		foreach ( $data_arr as $patient ) {
			$FIO                   = $patient['fio'];
			$birthday              = $patient['birthday'];
			$policy                = $patient['policy'];
			$sector                = $patient['sector'];
			$address               = $patient['address'];
			$phone                 = $patient['phone'];
			$fio_arr               = explode( ' ', $FIO );
			$data_json             = array();
			$data_json['Policy']   = $policy;
			$data_json['Surname']  = $fio_arr[0];
			$data_json['Name']     = $fio_arr[1];
			$data_json['Midname']  = $fio_arr[2];
			$data_json['phone']    = $phone;
			$data_json['birthday'] = $birthday;
			$data_json['address']  = $address;
			$data_json             = json_encode( $data_json, JSON_UNESCAPED_UNICODE );
			$query                 = "INSERT INTO wp_patients (FIO,Birthday)
				VALUES ('$FIO','$birthday')";
			$this->link->query( $query );
			$patient_id = $this->link->insert_id;


			//Проверяем, есть ли адрес в системе, есть ли есть - получаем id, если нет - добавляем адрес и добавляем связь - филиал-адрес
			$query       = "SELECT id FROM wp_address WHERE address='$address'";
			$address_obj = $this->link->query( $query );
			$row         = mysqli_fetch_assoc( $address_obj );
			$address_id  = $row['id'];


			if ( $address_id > 0 ) {

			}
			else {
				$query = "INSERT INTO wp_address (address)
							VALUES ('$address')";
				$this->link->query( $query );
				$address_id = $this->link->insert_id;
				$query      = "INSERT INTO wp_sector_address (sector,address_id)
						VALUES ($sector,$address_id)";
				$this->link->query( $query );


			}

			unset( $row );
			mysqli_free_result( $address_obj );

			if ( !$this->link->query( "CALL insertRelations($address_id, $patient_id, '$phone', '$data_json', '$policy')" ) ) {
				echo "349: Не удалось вызвать хранимую процедуру: (" . !$this->link->errno . ") " . !$this->link->error;
			}

			if ( $debug == 1 ) {
				//echo "<br>" . $query;
			}

		}
	}

	/**
	 * Обработка полиса
	 * @return mixed
	 */
	public function preparePolicy( $policy ) {

		return $policy = str_replace( ' ', '', trim( $policy ) );
	}

	/**
	 * Обработка даты для вставки в БД
	 * @return mixed
	 */
	public function prepareDate( $date ) {
		return implode( '-', array_reverse( explode( '.', $date ) ) );
	}

	/**
	 * Обработка телефона
	 * @return mixed
	 */
	public function preparePhone( $phone ) {
		$delimiter = ',';
		$result    = '';
		strpos( $phone, ';' ) != false ? $delimiter = ';' : $delimiter;
		$phones  = explode( $delimiter, $phone );
		$pattern = '/[^\d+]/';

		foreach ( $phones as $phone ) {
			$phone = preg_replace( $pattern, '', strval( $phone ) );


			switch ( strlen( $phone ) ) {
				case 7:
					$phone = substr( $phone, 0, 3 ) . '-' . substr( $phone, 3, 2 ) . '-' . substr( $phone, 5, 2 );
					break;
				case 11:
					$phone = $phone{0} . ' (' . substr( $phone, 1, 3 ) . ') ' . substr( $phone, 4, 3 ) . '-' . substr( $phone, 7, 2 ) . '-' . substr( $phone, 9, 2 );
					break;
				case 10:
					$phone = '8 (' . substr( $phone, 0, 3 ) . ') ' . substr( $phone, 3, 3 ) . '-' . substr( $phone, 6, 2 ) . '-' . substr( $phone, 8, 2 );
					break;
				default:
					strlen( $result ) > 0 ? $phone = "" : $phone = "Не указан";
					break;
			}
			$result .= $phone . ' ';
		}


		return trim( $result );
	}

	/**
	 * Очистка строки
	 * @return string
	 */
	public function escapeString( $string ) {
		return htmlspecialchars( strip_tags( trim( $string ) ) );
	}

	/**
	 * Очистска чисел
	 * @return int
	 */
	public function convertInt( $string ) {
		return (int)$string;
	}

}


$call_center = new CallCenterParser( $db_connect );
$data        = $call_center->getCitizens();
//$call_center->clearTables();
$call_center->insertData( $data, 1 );





