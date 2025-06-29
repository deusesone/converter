from openpyxl import load_workbook

data_file_path = 'inventory.xlsx'
wb = load_workbook(data_file_path)
ws = wb['proxmox']
row = list(ws.rows)
column = list(ws.columns)

output_file = open ('inventory.yaml', 'w')

for i in range (1, ws.max_row + 1):
    if i == 1:
        continue
    hostname = ws.cell(row=i, column=1).value
    print(hostname)
    ip = ws.cell(row=i, column=2).value
    print(ip)
    vm_id = ws.cell(row=i, column=3).value
    print(vm_id)
    bridge = ws.cell(row=i, column=4).value
    print(bridge)
    short_net = ws.cell(row=i, column=5).value
    print(short_net)
    vm_storage = ws.cell(row=i, column=6).value
    print(vm_storage)
    vm = ws.cell(row=i, column=7).value
    print(vm)
    template_id = ws.cell(row=i, column=8).value
    print(template_id)
    cores = ws.cell(row=i, column=9).value
    print(cores)
    memory = ws.cell(row=i, column=10).value
    print(memory)
    gateway = ws.cell(row=i, column=11).value
    print(gateway)
    proxmox_host = ws.cell(row=i, column=12).value
    print(proxmox_host)
    proxmox_ip = ws.cell(row=i, column=13).value
    print(proxmox_ip)
    print(hostname + ' ansible_host=' + ip + ' proxmox_vm_id=' + str(vm_id) + ' proxmox_vm_net_bridge=' + bridge + ' ansible_short_net=' + str(short_net) + ' proxmox_vm_storage=' + vm_storage + ' vm=' + vm + ' proxmox_template_id=' + str(template_id) + ' proxmox_vm_cores=' + str(cores) + ' proxmox_vm_mem=' + str(memory) + ' ct_gw=' + str(gateway) + ' proxmox_vm_host=' + proxmox_host + ' proxmox_delegate_host=' + proxmox_ip, file=output_file)