<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
exit();
/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/PHPExcel/IOFactory.php';
require_once dirname(__FILE__) . '/wp-config.php';
$dsn = 'mysql:dbname='.DB_NAME.';host=' . DB_HOST;
$user = DB_USER;
$password = DB_PASSWORD;

try {
	$dbh = new PDO($dsn, $user, $password);

} catch (PDOException $e) {
	echo 'Connection failed: ' . $e->getMessage();
}

$objPHPExcel = PHPExcel_IOFactory::load("shipping_import.xlsx");
$method_id = 11;
$count = 0;
$sheet = $objPHPExcel->setActiveSheetIndex($method_id)->toArray();
for($start = 2; $start<count($sheet); $start++) {
	$sth = "INSERT INTO `wp_woocommerce_shipping_table_rates` (`rate_class`,`rate_condition`,`rate_min`,`rate_max`,`rate_cost`,`rate_cost_per_item`,`rate_cost_per_weight_unit`,`rate_cost_percent`,`rate_label`,`rate_priority`,`rate_order`,`shipping_method_id`,`rate_abort`,`rate_abort_reason`) VALUE (:rate_class,:rate_condition,:rate_min,:rate_max,:rate_cost,:rate_cost_per_item,:rate_cost_per_weight_unit,:rate_cost_percent,:rate_label,:rate_priority,:rate_order,:shipping_method_id,:rate_abort,:rate_abort_reason)";

	$stmt = $dbh->prepare($sth);
	$data = array(
		'rate_class' => '',
		'rate_condition' => 'weight',
		'rate_min' => $sheet[$start][0],
		'rate_max' => $sheet[$start][1],
		'rate_cost' => $sheet[$start][2],
		'rate_cost_per_item' => 0,
		'rate_cost_per_weight_unit' => 0,
		'rate_cost_percent' => 0,
		'rate_label' => '',
		'rate_priority' => 0,
		'rate_order' => $count,
		'shipping_method_id' => $method_id+1,
		'rate_abort' => 0,
		'rate_abort_reason' => '',
	);
	$stmt->execute($data);
	$count++;

}
$dbh = null;
echo 'OK';exit();