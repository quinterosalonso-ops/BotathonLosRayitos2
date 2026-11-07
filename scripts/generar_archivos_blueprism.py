#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Empaqueta los objetos y procesos Blue Prism en un release importable."""

import os
import re
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BP_DIR = os.path.join(BASE_DIR, 'blueprism')


def component(filename):
    path = os.path.join(BP_DIR, filename)
    with open(path, 'r', encoding='utf-8-sig') as source:
        content = source.read()
    return re.sub(r'^\s*<\?xml[^>]*\?>\s*', '', content).strip()


def build_release():
    processes = [
        ('11111111-aaaa-4bbb-8ccc-000000000001', '01_Carga_Base_Contratos_BUK',
         component('03_BPA_Process_01_Carga_Base_Contratos_BUK.xml')),
        ('22222222-aaaa-4bbb-8ccc-000000000002', '02_Carga_y_Calculo_Planificacion',
         component('04_BPA_Process_02_Carga_y_Calculo_Planificacion.xml')),
        ('33333333-aaaa-4bbb-8ccc-000000000003', '03_Novedades_BUK_y_Contraste_SIC',
         component('05_BPA_Process_03_Novedades_BUK_y_Contraste_SIC.xml')),
    ]
    excel = component('01_BPA_Object_MS_Excel_VBO_Extended.xml')
    database = component('02_BPA_Object_Data_MSSQLLocalDB.xml')

    process_entries = ''.join(
        f'    <process id="{identifier}" name="{name}" xmlns="http://www.blueprism.co.uk/product/process">\n'
        f'      {content}\n'
        f'    </process>\n'
        for identifier, name, content in processes
    )
    release = f'''<?xml version="1.0" encoding="utf-8"?>
<bpr:release xmlns:bpr="http://www.blueprism.co.uk/product/release">
  <bpr:name>Botathon Teletón 2026 - Cada Hora Cuenta (MSSQLLocalDB + Excel Extended)</bpr:name>
  <bpr:release-notes>Paquete con MS Excel VBO - Extended, Data - MSSQLLocalDB y tres procesos de ingesta/cálculo/contraste para SQL Server LocalDB.</bpr:release-notes>
  <bpr:created>2026-11-07 05:00:00Z</bpr:created>
  <bpr:package-id>101</bpr:package-id>
  <bpr:package-name>Botathon_Teleton_2026_Completo</bpr:package-name>
  <bpr:user-created-by>admin</bpr:user-created-by>
  <bpr:contents count="7">
{process_entries}    <object id="30a1808c-9709-44d8-814a-1f17e2eb766f" name="MS Excel VBO - Extended" xmlns="http://www.blueprism.co.uk/product/process">
      {excel}
    </object>
    <object id="d91164ca-ea87-4f73-bad2-ee9010d382d9" name="Data - MSSQLLocalDB" xmlns="http://www.blueprism.co.uk/product/process">
      {database}
    </object>
    <process-group id="d133468a-5c27-4e11-a24e-65b32f51d832" name="Default" isDefaultGroup="True" xmlns="http://www.blueprism.co.uk/product/process-group">
      <members>
        <process id="11111111-aaaa-4bbb-8ccc-000000000001" />
        <process id="22222222-aaaa-4bbb-8ccc-000000000002" />
        <process id="33333333-aaaa-4bbb-8ccc-000000000003" />
      </members>
    </process-group>
    <object-group id="52f63a1e-b0a7-458f-b597-b603604eefc0" name="Default" isDefaultGroup="True" xmlns="http://www.blueprism.co.uk/product/object-group">
      <members>
        <object id="30a1808c-9709-44d8-814a-1f17e2eb766f" />
        <object id="d91164ca-ea87-4f73-bad2-ee9010d382d9" />
      </members>
    </object-group>
  </bpr:contents>
</bpr:release>'''

    output = os.path.join(BP_DIR, 'Botathon_Teleton_2026_Completo.bprelease')
    with open(output, 'w', encoding='utf-8') as target:
        target.write(release)
    ET.parse(output)
    print(f'OK -> Release generado y validado: {output}')


if __name__ == '__main__':
    build_release()
