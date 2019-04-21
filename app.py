from extended_minidom import create_tag, create_var
from datetime import datetime

doc = create_tag()
project = create_tag('project', attributes={'xmlns': 'http://www.plcopen.org/xml/tc6_0200'})
fileHeader = create_tag('fileHeader', attributes={'companyName': '',
                                                  'productName': 'CODESYS',
                                                  'productVersion': 'CODESYS V3.5 SP7 Patch 4',
                                                  'creationDateTime': str(datetime.now())})
contentHeader = create_tag('contentHeader', attributes={'name': 'kp15_20190419.project',
                                                        'modificationDateTime': str(datetime.now())})

coordinateInfo = create_tag('coordinateInfo')
fbd = create_tag('fbd')
ld = create_tag('ld')
sfc = create_tag('sfc')
scaling_fbd = create_tag('scaling', attributes={'x': '1', 'y': '1'})
scaling_ld = create_tag('scaling', attributes={'x': '1', 'y': '1'})
scaling_sfc = create_tag('scaling', attributes={'x': '1', 'y': '1'})
addData_ProjectInformation = create_tag('addData')
data_ProjectInformation = create_tag('data',
                                     attributes={'name': 'http://www.3s-software.com/plcopenxml/projectinformation',
                                                 'handleUnknown': 'implementation'})
ProjectInformation = create_tag('ProjectInformation')
types = create_tag('types')
dataTypes = create_tag('dataTypes')
pous = create_tag('pous')
pou = create_tag('pou', attributes={'name': 'POU',
                                    'pouType': 'program'})
interface = create_tag('interface')
localVars = create_tag('localVars')

doc.appendChild(project)
project.appendChild(fileHeader)
project.appendChild(contentHeader)
contentHeader.appendChild(coordinateInfo)
coordinateInfo.appendChild(fbd)
fbd.appendChild(scaling_fbd)
coordinateInfo.appendChild(ld)
ld.appendChild(scaling_ld)
coordinateInfo.appendChild(sfc)
sfc.appendChild(scaling_sfc)
contentHeader.appendChild(addData_ProjectInformation)
addData_ProjectInformation.appendChild(data_ProjectInformation)
data_ProjectInformation.appendChild(ProjectInformation)
project.appendChild(types)
types.appendChild(dataTypes)
types.appendChild(pous)
pous.appendChild(pou)

pou.appendChild(interface)
interface.appendChild(localVars)
create_var(localVars, var_name='state', var_type='INT ', var_simple_value='0')
create_var(localVars, var_name='i', var_type='WORD ')
create_var(localVars, var_name='j', var_type='WORD ')
create_var(localVars, var_name='res_fram_write', var_type='INT ', comment='результат записи 0 - всё хорошо			')
create_var(localVars, var_name='crc_fram_read', var_type='INT ', comment='посчитанный CRC из памяти')
create_var(localVars, var_name='srv_104', var_type='RLTNU.IEC104Server ', derived=True, comment=' IEC 104')
create_var(localVars, var_name='db_104', var_type='RLTNU.IECDB ', derived=True)
create_var(localVars, var_name='ts_104', var_type='RLTNU.IEC_M_SP_NA ', derived=True, array=['1', 'q_ts_full'],
           comment='мой массив бит на передачу в 104')
create_var(localVars, var_name='tf_104', var_type='RLTNU.IEC_M_ME_NC ', derived=True, array=['1', '(2 * q_ti_full)'],
           comment='мой массив float на передачу в 104')
create_var(localVars, var_name='tu_104', var_type='RLTNU.IECCommand ', derived=True, array=['1', 'q_tu_full'],
           comment=' массив команд в 104')
create_var(localVars, var_name='trf_104', var_type='RLTNU.IECCommand ', derived=True, array=['1', 'q_tr_full'],
           comment=' массив ТР в 104')
create_var(localVars, var_name='cmd_104', var_type='RLTNU.IECCommand ', derived=True)
create_var(localVars, var_name='event', var_type='RLTNU.IECSpontaneousEvent ', derived=True)
create_var(localVars, var_name='sys_tick', var_type='CAA.TICK ', derived=True, comment=' Время')
create_var(localVars, var_name='sys_tick_bool', var_type='BOOL ')
create_var(localVars, var_name='current_time', var_type='UDINT ')
create_var(localVars, var_name='hiTime', var_type='ULINT ')
create_var(localVars, var_name='Start_timer', var_type='TON', derived=True, comment=' Таймер первого запуска')
create_var(localVars, var_name='First_start', var_type='BOOL ', comment=' Флаг первого запуска')

body = create_tag('body')
ST = create_tag('ST')
xhtml = create_tag('xhtml', attributes={'xmlns': 'http://www.w3.org/1999/xhtml'})
ST_code = '''// Begin
sys_tick := TICKS.GetTick(sys_tick_bool);
current_time := CAA.TICK_TO_UDINT(sys_tick);
SysTimeRtc.SysTimeRtcHighResGet(hiTime);

CASE state OF
	0:
		////Настройка IEC104	//
		POU.IEC104_Init();

		////Работа с ТУ	//
		POU.TU_Init();
		
		////Работа с ТС//
		POU.TS_Init();
		
		////Работа с ТИ//		
		POU.TI_Init();
		
		First_start := TRUE;				// Поднимаем флаг первого запуска	
		state := 200;
		
	200:


		////Первый запуск//			
		Start_Timer( IN:=TRUE, PT:=T#5S ); // Запуск Таймера первого запуска

		
		////Запуск передачи данных IEC104//		
		srv_104();
		IF srv_104.error THEN
			state := 1000;
		END_IF
		
		/////Обработка и передача ТC
		POU.TS_Work(); 
		
		/////Обработка и передача ТИ
		POU.TI_Work();
		
		//// Индикация на шкафу
		POU.Valve_Indication();
		POU.Pump_Indication();
		
		///Обработка ТУ и ТР
		POU.TU_TR_WORK();

		
		
		ModbusMaster_SPort.enable := TRUE;	
		ModbusMaster_SPort();
		ModbusMaster_VPort.enable := TRUE;	
		ModbusMaster_VPort();
		
		/////Сброс флага первого запуска
		IF Start_Timer.Q THEN
			First_start := FALSE;
		END_IF
		
		state := 200;
	
	1000:
		srv_104.reset();
		state := 0;
		
END_CASE

// End'''

xml_str = doc.toprettyxml(indent="  ", encoding='utf-8')
with open("minidom_example2.xml", "wb") as f:
    f.write(xml_str)
