// SECTION Create any missing indexes and constraints
CREATE CONSTRAINT ON (vc:Vcenterserver) ASSERT vc.uid IS UNIQUE;
CREATE CONSTRAINT ON (vc:Vcenterserver) ASSERT vc.uid IS UNIQUE;
CREATE INDEX ON :Vcenterserver(name);
CREATE INDEX ON :Vcentercluster(name);
CREATE INDEX ON :Vcentercluster(managedby);
CREATE INDEX ON :Vspheredatacenter(name);
CREATE INDEX ON :Vspheredatacenter(managedby);
CREATE INDEX ON :Vresourcepool(name);
CREATE INDEX ON :Vspherehost(name);
CREATE INDEX ON :Vspherehost(objid);
CREATE INDEX ON :Vswitch(name);
CREATE INDEX ON :Vswitch(host);
CREATE INDEX ON :Virtualmachine(uuid);
CREATE INDEX ON :Virtualmachine(managedby);
CREATE INDEX ON :Vdatastore(name);
CREATE INDEX ON :Vdatastore(managedby);
CREATE INDEX ON :Vhostportgroup(name);
CREATE INDEX ON :Vhostportgroup(host);
CREATE INDEX ON :Vhostportgroup(managedby);


// SECTION Add .unverified property for any existing top-level nodes managed by this vCenter
// We will remove the property as we add them back in, so we can identify orphaned nodes
// Generally indicates these nodes were deleted/changed since the last time this import script was run
CALL apoc.load.xls("path-to-vmware-import-file",'vCluster',{header:true}) yield map as row
MATCH (vc:Vcenterserver {uid:row.`VI SDK UUID`}) 
MATCH (n) where n.managedby=vc.uid
SET n.unverified=true WITH n
OPTIONAL MATCH (n)-[r]-()
DELETE r;

// SECTION Add vCenter and clusters
CALL apoc.load.xls("path-to-vmware-import-file",'vCluster',{header:true}) yield map as row
MERGE (vc:Vcenterserver {uid:row.`VI SDK UUID`}) SET vc.name=row.`VI SDK Server`
MERGE (vrp:Vresourcepool {path:'None Configured',name:'None Configured',vc:row.`VI SDK Server`}) REMOVE vrp.unverified
MERGE (vpg:Vmportgroup {name:'None Provided',managedby:row.`VI SDK UUID`}) REMOVE vrp.unverified
MERGE (vcc:Vcentercluster {name:row.Name,managedby:row.`VI SDK UUID`}) REMOVE vcc.unverified
set vcc.hosts=row.OverallStatus,vcc.cpu=row.TotalCpu,vcc.CpuCored=row.NumCpuCores,vcc.memory=row.TotalMemory,vcc.ha=row.`HA enabled`,
vcc.drs=row.`DRS enabled`
MERGE (vcc)-[:CONTROLLED_BY_VC]-(vc);

// SECTION Set vCenter version & Build
CALL apoc.load.xls("path-to-vmware-import-file",'vInfo',{header:true}) yield map as row
WITH DISTINCT row.`VI SDK Server type` as vcversion,row.`VI SDK Server` as vcserver
WITH vcserver,split(vcversion,' build-')[0] as name,split(vcversion,' build-')[1] as build
MATCH (vc:Vcenterserver {name:vcserver})
MERGE (vcv:Vcenterversion {name:name})
MERGE (vcb:Vcenterbuild {build:build})
MERGE (vcb)-[:BUILD_OF]->(vcv)
MERGE (vc)-[:IS_VCENTER_BUILD]->(vcb);

// SECTION Create ():Vresourcepool) nodes, associate with datacenter, cluster, and vcenter
RETURN  "Create ():Vresourcepool) nodes, associate with datacenter, cluster, and vcenter...";
CALL apoc.load.xls("path-to-vmware-import-file",'vRP',{header:true}) yield map as row
with split(row.`Resource pool`,'Resources') as rp,row
with rp[0] as dcvmc,rp,row
with split(dcvmc,'/')[1] as datacenter,split(dcvmc,'/')[2] as cluster,row,rp[1] as resourcepool
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}), (vcc:Vcentercluster {name:cluster,managedby:row.`VI SDK UUID`})
MERGE (vdc:Vspheredatacenter {name:datacenter,managedby:row.`VI SDK UUID`})  REMOVE vdc.unverified
MERGE (vcc)-[:LOCATED_IN_DC]->(vdc)
MERGE (vdc)-[:CONTROLLED_BY_VC]->(vc)
with last(split(resourcepool,'/')) as pool,resourcepool,vc,vcc,vdc,row
with replace(resourcepool,'/'+pool,'') as parentpath,pool,vc,vcc,vdc,row
with parentpath,pool,vcc,vdc,vc,row, last(split(parentpath,'/')) as parent where pool <> ""
MERGE (vrp:Vresourcepool {name:pool,cluster:vcc.name,dc:vdc.name,vc:vc.name})
set vrp.vms=row.`# VMs`,vrp.cpus=row.`# vCPUs`,vrp.memcfg=row.`Mem Configured`,vrp.path=row.`Resource pool`
MERGE (vrp)-[:MEMBER_OF_CLUSTER]->(vcc)
WITH vrp,vc,parent,vcc,vdc
MATCH (pvrp:Vresourcepool {name:parent,cluster:vcc.name,dc:vdc.name,vc:vc.name})
MERGE (pvrp)<-[:CHILD_RESOURCE_POOL]-(vrp);

CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}), (vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`})
MERGE (vmh:Vspherehost {objid:row.`Object ID`,managedby:row.`VI SDK UUID`}) REMOVE vmh.unverified
MERGE (vmh)-[:CONTROLLED_BY_VC]-(vc)
MERGE (vmh)-[:MEMBER_OF_CLUSTER]->(vcc)
MERGE (cs:Vconfigstatus {name:row.`Config status`})
MERGE (vmh)-[:CONFIG_STATUS]->(cs)
SET vmh.name=row.Host,vmh.hosts=row.NumHosts,vmh.cpu=row.`# CPU`,vmh.cores=row.`# Cores`,vmh.memory=row.`# Memory`,vmh.memusage=row.`Memory usage %`,vmh.vms=row.`# VMs`
SET vmh.license=row.`Assigned License(s)`,vmh.chipset=row.`Max EVC`,vmh.boot=row.`Boot time`,vmh.servicetag=row.`Service tag`
MERGE (vpmp:Vspherecpupwrmgpol {name:row.`Current CPU power man. policy`})
MERGE (vmh)-[:IN_CPU_POW_MGMT]->(vpmp)
MERGE (vhpmp:Vspherehostpwrmgpol {name:row.`Host Power Policy`})
MERGE (vmh)-[:IN_HOST_POW_PLCY]->(vhpmp)
MERGE (cpum:Cpumodel {name:row.`CPU Model`})
MERGE (vmh)-[:HAS_CPU]->(cpum)
MERGE (vmhv:Vsphereesxversion {name:split(row.`ESX Version`,' build-')[0]})
MERGE (vmhb:Vsphereesxbuild {build:split(row.`ESX Version`,' build-')[1]})
MERGE (vmhb)-[:BUILD_OF]->(vmhv)
MERGE (vmh)-[:IS_ESX_BUILD]->(vmhb)
MERGE (vmh)-[:IS_ESX_VERSION]->(vmhv)
MERGE (cmfr:Crmmanufacturer {name:coalesce(row.Vendor,'None Provided')})
MERGE (vmh)-[:MANUFACTURED_BY]->(cmfr)
MERGE (m:Crmmodel {name:coalesce(row.Model,'None Provided')})
MERGE (vmh)-[:ASSET_MODEL]->(m)
MERGE (b:Biosversion {version:coalesce(row.`BIOS Version`,'None Provided'),date:row.`BIOS Date`})
MERGE (b)-[:MANUFACTURED_BY]->(cmfr)
MERGE (vmh)-[:BIOS_VERSION]->(b)
with vmh,row
MATCH (cd:Clientdomain {name:coalesce(row.Domain,'None Provided')})--(a:Company)
MERGE (vmh)-[:OF_DOMAIN]->(cd)
MERGE (vmh)-[:ESX_HOST_FOR]->(a);

// SECTION  Create (:Ntpserver) nodes (by IP) and RELATIONSHIP (:Vspherehost)-[:USES_NTP]->(:Ntpserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`NTP Server(s)`,',') as ntpservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND ntpservers as ntph
WITH vmh,trim(ntph) as ntp where trim(ntph)=~regex
MERGE (ntph:Ntpserver {ipaddress:ntp})
MERGE (vmh)-[:USES_NTP]->(ntph);

// SECTION  Create (:Ntpserver) nodes (by fqdn) and RELATIONSHIP (:Vspherehost)-[:USES_NTP]->(:Ntpserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`NTP Server(s)`,',') as ntpservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND ntpservers as ntph
WITH vmh,trim(ntph) as ntp where not(trim(ntph)=~regex)
MERGE (ntph:Ntpserver {fqdn:ntp})
MERGE (vmh)-[:USES_NTP]->(ntph);

// SECTION  Create (:Dnsserver) nodes (by IP) and RELATIONSHIP (:Vspherehost)-[:USES_DNS]->(:Dnsserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`DNS Servers`,',') as dnsservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND dnsservers as dnsh
WITH vmh,trim(dnsh) as dns where trim(dnsh)=~regex
MERGE (dnss:Dnsserver {ipaddress:dns})
MERGE (vmh)-[:USES_DNS]->(dnss);

// SECTION  Create (:Dnsserver) nodes (by IP) and RELATIONSHIP (:Vspherehost)-[:USES_DNS]->(:Dnsserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`DNS Servers`,',') as dnsservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND dnsservers as dnsh
WITH vmh,trim(dnsh) as dns where not(trim(dnsh)=~regex)
MERGE (dnss:Dnsserver {fqdn:dns})
MERGE (vmh)-[:USES_DNS]->(dnss);

// SECTION  Create (:Vswitch) and RELATIONSHIPS to (:Vspherehost), (:Vswlbpolicy), and (:Jumboframes)
CALL apoc.load.xls("path-to-vmware-import-file",'vSwitch',{header:true}) yield map as row
MATCH (vmh:Vspherehost {name:row.Host})--(vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`})
MERGE (vsw:Vswitch {name:row.Switch,host:row.Host})
SET vsw.ports=row.`# Ports`,vsw.freeports=row.`Free Ports`,vsw.promiscuous=row.`Promiscuous Mode`,vsw.macchanges=row.`Mac Changes`,vsw.forged=row.`Forged Transmits`
SET vsw.shaping=row.`Traffic Shaping`,vsw.notifysw=row.`Notify Switch`,vsw.mtu=toInt(row.MTU),vsw.offload=row.Offload
MERGE (vsw)-[:VSWITCH_FOR_HOST]->(vmh)
MERGE (vsp:Vlbpolicy {name:row.Policy})
MERGE (vsw)-[:LOAD_BALANCING_POLICY]->(vsp)
WITH vsw
MATCH (jmb:Jumboframes {name:'enabled'}),(vsw) where vsw.mtu >= 9000
MERGE (vsw)-[:HAS_JUMBO_FRAMES]->(jmb);

// SECTION  Create (:Vportgroup) and RELATIONSHIPS to (:Vspherehost), (:Vswlbpolicy), and (:Jumboframes)
CALL apoc.load.xls("path-to-vmware-import-file",'vPort',{header:true}) yield map as row
MATCH (vmh:Vspherehost {name:row.Host})--(vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`}),(vsw:Vswitch {name:row.Switch,host:row.Host})
MERGE (vpg:Vportgroup {name:row.`Port Group`,managedby:row.`VI SDK UUID`}) REMOVE vpg.unverified
MERGE (pg:Vhostportgroup {name:row.`Port Group`,host:row.Host,managedby:row.`VI SDK UUID`}) REMOVE pg.unverified
MERGE (vsp:Vlbpolicy {name:coalesce(row.Policy,'None Provided')})
MERGE (vpg)<-[:HOST_PG_FOR]-(pg)
MERGE (pg)-[:STANDARD_PG_ON]->(vmh)
MERGE (vsw)-[:LOAD_BALANCING_POLICY]->(vsp)
SET pg.vlan=row.VLAN,pg.promiscuous=row.`Promiscuous Mode`,pg.macchanges=row.`Mac Changes`,pg.forged=row.`Forged Transmits`,pg.shaping=row.`Traffic Shaping`;

// SECTION  Create (:Vmnic) and RELATIONSHIPS  from vNIC tab
CALL apoc.load.xls("path-to-vmware-import-file",'vNIC',{header:true}) yield map as row
with row,coalesce(row.Speed,'No link') as linkspeed,coalesce(row.Driver,'None Provided') as nicdriver
MATCH (vmh:Vspherehost {name:row.Host})--(vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`}),(vsw:Vswitch {name:row.Switch,host:row.Host})
MERGE (vmnic:Vmnic {name:row.`Network Device`,host:row.Host})
MERGE (vnd:Vmnicdriver {name:nicdriver})
MERGE (vmnic)-[:USES_DRIVER]->(vnd)
MERGE (vns:Vmnicspeed {name:linkspeed})
MERGE (vmnic)-[:LINK_SPEED]-(vns)
MERGE (vmnic)-[:PNIC_OF_HOST]-(vmh)
SET vmnic.mac=row.MAC,vmnic.wake=row.WakeOn,vmnic.pci=row.PCI
MERGE (vmnic)<-[:NETWORK_ADAPTERS]-(vsw);

// SECTION  MERGE (:Virtualmachine) from vInfo tab
CALL apoc.load.xls("path-to-vmware-import-file",'vInfo',{header:true}) yield map as row
OPTIONAL MATCH (vdc:Vspheredatacenter {name:split(row.`Folder`,'/')[1],managedby:row.`VI SDK UUID`}) 
OPTIONAL MATCH (vcc:Vcentercluster {name:split(row.`Resource pool`,'/')[2],managedby:row.`VI SDK UUID`})
MERGE (vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`}) REMOVE vm.unverified
SET vm.name=row.VM,vm.fqdn=vm.`DNS Name`,vm.poweron=row.`PowerOn`,vm.changedon=row.`Change Version`,vm.note=row.Annotation,vm.vmid=row.`VM ID`
SET vm.needsconsolidation=row.`Consolidation Needed`,vm.cpus=row.CPUs,vm.memory=toInt(row.Memory),vm.nics=toInt(row.NICs),vm.disks=toInt(row.Disks),vm.cbt=row.CBT
MERGE (vcpu:Vcpus {name:row.CPUs + ' vCPUs',qty:toInt(row.CPUs)})
MERGE (vhwv:Vhwver {name:toInt(row.`HW version`)})
MERGE (vm)-[hvr:HW_VERSION]->(vhwv)
MERGE (vm)-[:HAS_VCPUS]->(vcpu)
MERGE (cs:Vconnectionstate {name:row.`Connection state`})
MERGE (vm)-[:CONNECTION_STATE]->(cs)
MERGE (cfgs:Vconfigstatus {name:row.`Config status`})
MERGE (vm)-[:CONFIG_STATUS]->(cfgs)
MERGE (vmps:Vmpwrstate {name:row.Powerstate})
MERGE (vm)-[:IN_POWER_STATE]->(vmps)
MERGE (vmgs:Vmpgueststate {name:row.`Guest state`})
MERGE (vm)-[:IN_GUEST_STATE]->(vmgs)
MERGE (vmhb:Vmheartbeat {name:row.Heartbeat})
MERGE (vm)-[:HEARTBEAT]->(vmhb)
FOREACH (ignoreMe in CASE WHEN exists(row.`Resource pool`) and length(split(row.`Resource pool`,'/'))>4 THEN [1] ELSE [] END | MERGE (vrp:Vresourcepool {path:coalesce(row.`Resource pool`,'None Configured'),vc:row.`VI SDK Server`}) SET vrp.name =last(split(row.`Resource pool`,'/'))
MERGE (vm)-[:IN_RESOURCE_POOL]->(vrp))
FOREACH (ignoreMe in CASE WHEN (exists(row.Folder) and length(split(row.Folder,'/'))>2)  THEN [1] ELSE [] END | MERGE (vfl:Vfolder {path:row.`Folder`}) MERGE (vm)-[:IN_FOLDER]->(vfl) SET vfl.name=last(split(row.`Folder`,'/')))
FOREACH (ignoreMe in CASE WHEN (exists(row.`OS according to the VMware Tools`)) THEN [1] ELSE [] END | MERGE (vmost:Vmos {name:coalesce(row.`OS according to the VMware Tools`,'None Provided')}) MERGE (vm)-[:OS_VIA_TOOLS]-(vmost))
FOREACH (ignoreMe in CASE WHEN (exists(row.`OS according to the configuration file`)) THEN [1] ELSE [] END | MERGE (vmos:Vmos {name:coalesce(row.`OS according to the configuration file`,'None Provided')}) MERGE (vm)-[:OS_VIA_CONFIG]-(vmos))
MERGE (vmpg1:Vportgroup {name:coalesce(row.`Network #1`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg1)
MERGE (vmpg2:Vportgroup {name:coalesce(row.`Network #2`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg2)
MERGE (vmpg3:Vportgroup {name:coalesce(row.`Network #3`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg3)
MERGE (vmpg4:Vportgroup {name:coalesce(row.`Network #4`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg4)
SET hvr.upgradestatus=row.`HW upgrade status`
WITH *
OPTIONAL MATCH (vfl:Vfolder {path:row.`Folder`})
OPTIONAL MATCH (vrp:Vresourcepool {path:row.`Resource pool`})
OPTIONAL MATCH (pvfl:Vfolder {path:replace(vfl.path,'/'+vfl.name,'')})
OPTIONAL MATCH (pvrp:Vresourcepool {path:replace(vrp.path,'/'+vrp.name,'')})
FOREACH (ignoreMe in CASE WHEN exists(vfl.path) and exists(pvfl.path) THEN [1] ELSE [] END | MERGE (vfl)-[:IN_FOLDER]->(pvfl))
FOREACH (ignoreMe in CASE WHEN exists(vfl.path) and pvfl.path is null THEN [1] ELSE [] END | MERGE (vfl)-[:LOCATED_IN_DC]->(vdc))
FOREACH (ignoreMe in CASE WHEN vfl.path is null THEN [1] ELSE [] END | MERGE (vm)-[:LOCATED_IN_DC]->(vdc))
FOREACH (ignoreMe in CASE WHEN exists(vrp.path) and exists(pvrp.path) THEN [1] ELSE [] END | MERGE (vrp)-[:CHILD_RESOURCE_OF]->(pvrp))
FOREACH (ignoreMe in CASE WHEN exists(vcc.name) and exists(vrp.path) and pvrp.path is null THEN [1] ELSE [] END | MERGE (vrp)-[:LOCATED_IN_CLUSTER]->(vcc))
FOREACH (ignoreMe in CASE WHEN exists(vcc.name) and vrp.path is null THEN [1] ELSE [] END | MERGE (vm)-[:LOCATED_IN_CLUSTER]->(vcc))
RETURN count(vm) AS `Creating (:Virtualmachine) nodes and relationships to their state...`;


// SECTION  MERGE (:Vdatastore) from vDatastore tab
CALL apoc.load.xls("path-to-vmware-import-file",'vDatastore',{header:true}) yield map as row
MATCH (vc:Vcenterserver {uid:row.`VI SDK UUID`})
MERGE (ds:Vdatastore {url:row.URL}) REMOVE ds.unverified
SET ds.name=row.Name,ds.accessible=row.Accessible,ds.capacity=row.`Capacity MB`,ds.inuse=row.`In Use MB`,ds.free=row.`Free MB`
SET ds.hosts=row.`# Hosts`,ds.verion=row.Version,ds.sio=row.`SIOC enabled`,ds.vms=row.`# VMs`,ds.address=row.Address,ds.managedby=row.`VI SDK UUID`
MERGE (cs:Vconfigstatus {name:row.`Config status`})
MERGE (ds)-[:CONFIG_STATUS]->(cs)
MERGE (vt:Vdatastoretype {name:row.Type})
MERGE (ds)-[:DATASTORE_TYPE]->(vt)
WITH row,ds,vc,split(row.Hosts,',') as hosts
UNWIND hosts as vmhost
MATCH (vmh:Vspherehost {name:trim(vmhost),managedby:row.`VI SDK UUID`})
MERGE (vmh)-[:CONNECTED_DATASTORE]->(ds);

// SECTION  MERGE (:Virtualdisk) from vDisk tab
CALL apoc.load.xls("path-to-vmware-import-file",'vDisk',{header:true}) yield map as row
MATCH (vc:Vcenterserver {uid:row.`VI SDK UUID`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vd:Virtualdisk {path:row.Path})
SET vd.disk=row.Disk,vd.capacity=row.`Capacity MB`,vd.thin=row.Thin,vd.controller=row.Controller
SET vd.mode=row.`Disk Mode`,vd.eager=row.`Eagerly Scrub`,vd.template=row.Template
MERGE (vd)-[:VDISK_FOR_VM]-(vm)
with row,vc,vd,replace(split(row.Path,']')[0],'[','') as dsname
MATCH (ds:Vdatastore {name:dsname, managedby:row.`VI SDK UUID`})--(:Vspherehost {name:row.Host,managedby:row.`VI SDK UUID`})
MERGE (vd)-[:ON_DATASTORE]-(ds);

// SECTION  MERGE (:Vmadapter) from vNetwork
CALL apoc.load.xls("path-to-vmware-import-file",'vNetwork',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vmn:Vmadapter {mac:row.`Mac Address`,vmuuid:row.`VM UUID`})
MERGE (vmn)-[:ADAPTER_FOR]-(vm)
MERGE (vmat:Vmadaptertype {name:row.Adapter})
MERGE (vmn)-[:ADAPTER_TYPE]-(vmat)
SET vmn.startconnected=row.`Starts Connected`,vmn.ip=row.`IP Address`
with row,vmn
MATCH (pg:Vhostportgroup {name:row.Network,host:row.Host,managedby:row.`VI SDK UUID`})
MERGE (vmn)-[:IN_PORTGROUP]->(pg);

// SECTION  MERGE (:Vpartition) from vPartition
CALL apoc.load.xls("path-to-vmware-import-file",'vPartition',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vmp:Vpartition {disk:row.Disk,vmuuid:row.`VM UUID`})
MERGE (vmp)-[:PARTITION_FOR]-(vm)
SET vmp.capacity=row.`Capacity MB`,vmp.consumed=row.`Consumed MB`,vmp.free=row.`Free %`;

// SECTION  MERGE (:Vpartition) from vPartition
CALL apoc.load.xls("path-to-vmware-import-file",'vSnapshot',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vmss:Vsnapshot {name:row.Name,vmuuid:row.`VM UUID`})
MERGE (vmss)-[:SNAPSHOT_OF]-(vm)
SET vmss.description=row.Description,vmss.timestamp=row.`Date / time`,vmss.size=row.`Size MB (total)`;

// SECTION  Add vCenter and clusters
CALL apoc.load.xls("path-to-vmware-import-file",'vCluster',{header:true}) yield map as row
MERGE (vc:Vcenterserver {name:row.`VI SDK Server`})
MERGE (vrp:Vresourcepool {path:'None Configured',name:'None Configured',vc:row.`VI SDK Server`}) REMOVE vrp.unverified
MERGE (vmpg:Vmportgroup {name:'None Provided',managedby:row.`VI SDK UUID`}) REMOVE vmpg.unverified
MERGE (vcc:Vcentercluster {name:row.Name,managedby:row.`VI SDK UUID`})
ON CREATE
set vcc.hosts=row.OverallStatus,vcc.cpu=row.TotalCpu,vcc.CpuCored=row.NumCpuCores,vcc.memory=row.TotalMemory,vcc.ha=row.`HA enabled`,
vcc.drs=row.`DRS enabled`
MERGE (vcc)-[:CONTROLLED_BY_VC]-(vc);

// SECTION  Set vCenter version & Build
CALL apoc.load.xls("path-to-vmware-import-file",'vInfo',{header:true}) yield map as row
WITH DISTINCT row.`VI SDK Server type` as vcversion,row.`VI SDK Server` as vcserver
WITH vcserver,split(vcversion,' build-')[0] as name,split(vcversion,' build-')[1] as build
MATCH (vc:Vcenterserver {name:vcserver})
MERGE (vcv:Vcenterversion {name:name})
MERGE (vcb:Vcenterbuild {build:build})
MERGE (vcb)-[:BUILD_OF]->(vcv)
MERGE (vc)-[:IS_VCENTER_BUILD]->(vcb);

// SECTION Create ():Vresourcepool) nodes, associate with datacenter, cluster, and vcenter
RETURN  "Create ():Vresourcepool) nodes, associate with datacenter, cluster, and vcenter...";
CALL apoc.load.xls("path-to-vmware-import-file",'vRP',{header:true}) yield map as row
with split(row.`Resource pool`,'Resources') as rp,row
with rp[0] as dcvmc,rp,row
with split(dcvmc,'/')[1] as datacenter,split(dcvmc,'/')[2] as cluster,row,rp[1] as resourcepool
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}), (vcc:Vcentercluster {name:cluster,managedby:row.`VI SDK UUID`})
MERGE (vdc:Vspheredatacenter {name:datacenter,managedby:row.`VI SDK UUID`})
MERGE (vcc)-[:LOCATED_IN_DC]->(vdc)
MERGE (vdc)-[:CONTROLLED_BY_VC]->(vc)
with last(split(resourcepool,'/')) as pool,resourcepool,vc,vcc,vdc,row
with replace(resourcepool,'/'+pool,'') as parentpath,pool,vc,vcc,vdc,row
with parentpath,pool,vcc,vdc,vc,row, last(split(parentpath,'/')) as parent where pool <> ""
MERGE (vrp:Vresourcepool {name:pool,cluster:vcc.name,dc:vdc.name,vc:vc.name})
set vrp.vms=row.`# VMs`,vrp.cpus=row.`# vCPUs`,vrp.memcfg=row.`Mem Configured`,vrp.path=row.`Resource pool`
MERGE (vrp)-[:MEMBER_OF_CLUSTER]->(vcc)
WITH vrp,vc,parent,vcc,vdc
MATCH (pvrp:Vresourcepool {name:parent,cluster:vcc.name,dc:vdc.name,vc:vc.name})
MERGE (pvrp)<-[:CHILD_RESOURCE_POOL]-(vrp);

CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}), (vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`})
MERGE (vmh:Vspherehost {objid:row.`Object ID`,managedby:row.`VI SDK UUID`})
MERGE (vmh)-[:CONTROLLED_BY_VC]-(vc)
MERGE (vmh)-[:MEMBER_OF_CLUSTER]->(vcc)
MERGE (cs:Vconfigstatus {name:row.`Config status`})
MERGE (vmh)-[:CONFIG_STATUS]->(cs)
SET vmh.name=row.Host,vmh.hosts=row.NumHosts,vmh.cpu=row.`# CPU`,vmh.cores=row.`# Cores`,vmh.memory=row.`# Memory`,vmh.memusage=row.`Memory usage %`,vmh.vms=row.`# VMs`
SET vmh.license=row.`Assigned License(s)`,vmh.chipset=row.`Max EVC`,vmh.boot=row.`Boot time`,vmh.servicetag=row.`Service tag`
MERGE (vpmp:Vspherecpupwrmgpol {name:row.`Current CPU power man. policy`})
MERGE (vmh)-[:IN_CPU_POW_MGMT]->(vpmp)
MERGE (vhpmp:Vspherehostpwrmgpol {name:row.`Host Power Policy`})
MERGE (vmh)-[:IN_HOST_POW_PLCY]->(vhpmp)
MERGE (cpum:Cpumodel {name:row.`CPU Model`})
MERGE (vmh)-[:HAS_CPU]->(cpum)
MERGE (vmhv:Vsphereesxversion {name:split(row.`ESX Version`,' build-')[0]})
MERGE (vmhb:Vsphereesxbuild {build:split(row.`ESX Version`,' build-')[1]})
MERGE (vmhb)-[:BUILD_OF]->(vmhv)
MERGE (vmh)-[:IS_ESX_BUILD]->(vmhb)
MERGE (vmh)-[:IN_DOMAIN]->(vmhv)
MERGE (cmfr:Crmmanufacturer {name:coalesce(row.Vendor,'None Provided')})
MERGE (vmh)-[:MANUFACTURED_BY]->(cmfr)
MERGE (m:Crmmodel {name:coalesce(row.Model,'None Provided')})
MERGE (vmh)-[:ASSET_MODEL]->(m)
MERGE (b:Biosversion {version:coalesce(row.`BIOS Version`,'None Provided'),date:row.`BIOS Date`})
MERGE (b)-[:MANUFACTURED_BY]->(cmfr)
MERGE (vmh)-[:BIOS_VERSION]->(b)
with vmh,row
MATCH (cd:Clientdomain {name:coalesce(row.Domain,'None Provided')})--(a:Company)
MERGE (vmh)-[:OF_DOMAIN]->(cd)
MERGE (vmh)-[:ESX_HOST_FOR]->(a);

// SECTION  Create (:Ntpserver) nodes (by IP) and RELATIONSHIP (:Vspherehost)-[:USES_NTP]->(:Ntpserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`NTP Server(s)`,',') as ntpservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND ntpservers as ntph
WITH vmh,trim(ntph) as ntp where trim(ntph)=~regex
MERGE (ntph:Ntpserver {ipaddress:ntp})
MERGE (vmh)-[:USES_NTP]->(ntph);

// SECTION  Create (:Ntpserver) nodes (by fqdn) and RELATIONSHIP (:Vspherehost)-[:USES_NTP]->(:Ntpserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`NTP Server(s)`,',') as ntpservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND ntpservers as ntph
WITH vmh,trim(ntph) as ntp where not(trim(ntph)=~regex)
MERGE (ntph:Ntpserver {fqdn:ntp})
MERGE (vmh)-[:USES_NTP]->(ntph);

// SECTION  Create (:Dnsserver) nodes (by IP) and RELATIONSHIP (:Vspherehost)-[:USES_DNS]->(:Dnsserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`DNS Servers`,',') as dnsservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND dnsservers as dnsh
WITH vmh,trim(dnsh) as dns where trim(dnsh)=~regex
MERGE (dnss:Dnsserver {ipaddress:dns})
MERGE (vmh)-[:USES_DNS]->(dnss);

// SECTION  Create (:Dnsserver) nodes (by IP) and RELATIONSHIP (:Vspherehost)-[:USES_DNS]->(:Dnsserver)
CALL apoc.load.xls("path-to-vmware-import-file",'vHost',{header:true}) yield map as row
MATCH (vmh:Vspherehost {objid:row.`Object ID`,name:row.Host})
WITH vmh,split(row.`DNS Servers`,',') as dnsservers,'\\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b' as regex
UNWIND dnsservers as dnsh
WITH vmh,trim(dnsh) as dns where not(trim(dnsh)=~regex)
MERGE (dnss:Dnsserver {fqdn:dns})
MERGE (vmh)-[:USES_DNS]->(dnss);

// SECTION  Create (:Vswitch) and RELATIONSHIPS to (:Vspherehost), (:Vswlbpolicy), and (:Jumboframes)
CALL apoc.load.xls("path-to-vmware-import-file",'vSwitch',{header:true}) yield map as row
MATCH (vmh:Vspherehost {name:row.Host})--(vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`})
MERGE (vsw:Vswitch {name:row.Switch,host:row.Host})
SET vsw.ports=row.`# Ports`,vsw.freeports=row.`Free Ports`,vsw.promiscuous=row.`Promiscuous Mode`,vsw.macchanges=row.`Mac Changes`,vsw.forged=row.`Forged Transmits`
SET vsw.shaping=row.`Traffic Shaping`,vsw.notifysw=row.`Notify Switch`,vsw.mtu=toInt(row.MTU),vsw.offload=row.Offload
MERGE (vsw)-[:VSWITCH_FOR_HOST]->(vmh)
MERGE (vsp:Vlbpolicy {name:row.Policy})
MERGE (vsw)-[:LOAD_BALANCING_POLICY]->(vsp)
WITH vsw
MATCH (jmb:Jumboframes {name:'enabled'}),(vsw) where vsw.mtu >= 9000
MERGE (vsw)-[:HAS_JUMBO_FRAMES]->(jmb);

// SECTION  Create (:Vportgroup) and RELATIONSHIPS to (:Vspherehost), (:Vswlbpolicy), and (:Jumboframes)
CALL apoc.load.xls("path-to-vmware-import-file",'vPort',{header:true}) yield map as row
MATCH (vmh:Vspherehost {name:row.Host})--(vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`}),(vsw:Vswitch {name:row.Switch,host:row.Host})
MERGE (vpg:Vportgroup {name:row.`Port Group`,managedby:row.`VI SDK UUID`}) REMOVE vpg.unverified
MERGE (pg:Vhostportgroup {name:row.`Port Group`,host:row.Host,managedby:row.`VI SDK UUID`})
MERGE (vsp:Vlbpolicy {name:coalesce(row.Policy,'None Provided')})
MERGE (vpg)<-[:HOST_PG_FOR]-(pg)
MERGE (vsw)-[:LOAD_BALANCING_POLICY]->(vsp)
SET pg.vlan=row.VLAN,pg.promiscuous=row.`Promiscuous Mode`,pg.macchanges=row.`Mac Changes`,pg.forged=row.`Forged Transmits`,pg.shaping=row.`Traffic Shaping`;

// SECTION  Create (:Vmnic) and RELATIONSHIPS  from vNIC tab
CALL apoc.load.xls("path-to-vmware-import-file",'vNIC',{header:true}) yield map as row
with row,coalesce(row.Speed,'No link') as linkspeed,coalesce(row.Driver,'None Provided') as nicdriver
MATCH (vmh:Vspherehost {name:row.Host})--(vcc:Vcentercluster {name:row.Cluster,managedby:row.`VI SDK UUID`}),(vsw:Vswitch {name:row.Switch,host:row.Host})
MERGE (vmnic:Vmnic {name:row.`Network Device`,host:row.Host})
MERGE (vnd:Vmnicdriver {name:nicdriver})
MERGE (vmnic)-[:USES_DRIVER]->(vnd)
MERGE (vns:Vmnicspeed {name:linkspeed})
MERGE (vmnic)-[:LINK_SPEED]-(vns)
MERGE (vmnic)-[:PNIC_OF_HOST]-(vmh)
SET vmnic.mac=row.MAC,vmnic.wake=row.WakeOn,vmnic.pci=row.PCI
MERGE (vmnic)<-[:NETWORK_ADAPTERS]-(vsw);

// SECTION  MERGE (:Virtualmachine) from vInfo tab
CALL apoc.load.xls("path-to-vmware-import-file",'vInfo',{header:true}) yield map as row
OPTIONAL MATCH (vdc:Vspheredatacenter {name:split(row.`Folder`,'/')[1],managedby:row.`VI SDK UUID`}) 
OPTIONAL MATCH (vcc:Vcentercluster {name:split(row.`Resource pool`,'/')[2],managedby:row.`VI SDK UUID`})
MERGE (vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`}) REMOVE vm.unverified
SET vm.name=row.VM,vm.fqdn=vm.`DNS Name`,vm.poweron=row.`PowerOn`,vm.changedon=row.`Change Version`,vm.note=row.Annotation,vm.vmid=row.`VM ID`
SET vm.needsconsolidation=row.`Consolidation Needed`,vm.cpus=row.CPUs,vm.memory=toInt(row.Memory),vm.nics=toInt(row.NICs),vm.disks=toInt(row.Disks),vm.cbt=row.CBT
MERGE (vcpu:Vcpus {name:row.CPUs + ' vCPUs',qty:toInt(row.CPUs)})
MERGE (vhwv:Vhwver {name:toInt(row.`HW version`)})
MERGE (vm)-[hvr:HW_VERSION]->(vhwv)
MERGE (vm)-[:HAS_VCPUS]->(vcpu)
MERGE (cs:Vconnectionstate {name:row.`Connection state`})
MERGE (vm)-[:CONNECTION_STATE]->(cs)
MERGE (cfgs:Vconfigstatus {name:row.`Config status`})
MERGE (vm)-[:CONFIG_STATUS]->(cfgs)
MERGE (vmps:Vmpwrstate {name:row.Powerstate})
MERGE (vm)-[:IN_POWER_STATE]->(vmps)
MERGE (vmgs:Vmpgueststate {name:row.`Guest state`})
MERGE (vm)-[:IN_GUEST_STATE]->(vmgs)
MERGE (vmhb:Vmheartbeat {name:row.Heartbeat})
MERGE (vm)-[:HEARTBEAT]->(vmhb)
FOREACH (ignoreMe in CASE WHEN exists(row.`Resource pool`) and length(split(row.`Resource pool`,'/'))>4 THEN [1] ELSE [] END | MERGE (vrp:Vresourcepool {path:coalesce(row.`Resource pool`,'None Configured'),vc:row.`VI SDK Server`}) SET vrp.name =last(split(row.`Resource pool`,'/'))
MERGE (vm)-[:IN_RESOURCE_POOL]->(vrp))
FOREACH (ignoreMe in CASE WHEN (exists(row.Folder) and length(split(row.Folder,'/'))>2)  THEN [1] ELSE [] END | MERGE (vfl:Vfolder {path:row.`Folder`}) MERGE (vm)-[:IN_FOLDER]->(vfl) SET vfl.name=last(split(row.`Folder`,'/')))
FOREACH (ignoreMe in CASE WHEN (exists(row.`OS according to the VMware Tools`)) THEN [1] ELSE [] END | MERGE (vmost:Vmos {name:coalesce(row.`OS according to the VMware Tools`,'None Provided')}) MERGE (vm)-[:OS_VIA_TOOLS]-(vmost))
FOREACH (ignoreMe in CASE WHEN (exists(row.`OS according to the configuration file`)) THEN [1] ELSE [] END | MERGE (vmos:Vmos {name:coalesce(row.`OS according to the configuration file`,'None Provided')}) MERGE (vm)-[:OS_VIA_CONFIG]-(vmos))
MERGE (vmpg1:Vportgroup {name:coalesce(row.`Network #1`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg1)
MERGE (vmpg2:Vportgroup {name:coalesce(row.`Network #2`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg2)
MERGE (vmpg3:Vportgroup {name:coalesce(row.`Network #3`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg3)
MERGE (vmpg4:Vportgroup {name:coalesce(row.`Network #4`,'Not Configured'),managedby:row.`VI SDK UUID`})
MERGE (vm)-[:IN_PORTGROUP]->(vmpg4)
SET hvr.upgradestatus=row.`HW upgrade status`
WITH *
OPTIONAL MATCH (vfl:Vfolder {path:row.`Folder`})
OPTIONAL MATCH (vrp:Vresourcepool {path:row.`Resource pool`})
OPTIONAL MATCH (pvfl:Vfolder {path:replace(vfl.path,'/'+vfl.name,'')})
OPTIONAL MATCH (pvrp:Vresourcepool {path:replace(vrp.path,'/'+vrp.name,'')})
FOREACH (ignoreMe in CASE WHEN exists(vfl.path) and exists(pvfl.path) THEN [1] ELSE [] END | MERGE (vfl)-[:IN_FOLDER]->(pvfl))
FOREACH (ignoreMe in CASE WHEN exists(vfl.path) and pvfl.path is null THEN [1] ELSE [] END | MERGE (vfl)-[:LOCATED_IN_DC]->(vdc))
FOREACH (ignoreMe in CASE WHEN vfl.path is null THEN [1] ELSE [] END | MERGE (vm)-[:LOCATED_IN_DC]->(vdc))
FOREACH (ignoreMe in CASE WHEN exists(vrp.path) and exists(pvrp.path) THEN [1] ELSE [] END | MERGE (vrp)-[:CHILD_RESOURCE_OF]->(pvrp))
FOREACH (ignoreMe in CASE WHEN exists(vcc.name) and exists(vrp.path) and pvrp.path is null THEN [1] ELSE [] END | MERGE (vrp)-[:LOCATED_IN_CLUSTER]->(vcc))
FOREACH (ignoreMe in CASE WHEN exists(vcc.name) and vrp.path is null THEN [1] ELSE [] END | MERGE (vm)-[:LOCATED_IN_CLUSTER]->(vcc))
RETURN count(vm) AS `Creating (:Virtualmachine) nodes and relationships to their state...`;


// SECTION  MERGE (:Vdatastore) from vDatastore tab
CALL apoc.load.xls("path-to-vmware-import-file",'vDatastore',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`})
MERGE (ds:Vdatastore {url:row.URL})
SET ds.name=row.Name,ds.accessible=row.Accessible,ds.capacity=row.`Capacity MB`,ds.inuse=row.`In Use MB`,ds.free=row.`Free MB`
SET ds.hosts=row.`# Hosts`,ds.verion=row.Version,ds.sio=row.`SIOC enabled`,ds.vms=row.`# VMs`,ds.address=row.Address,ds.managedby=row.`VI SDK UUID`
MERGE (cs:Vconfigstatus {name:row.`Config status`})
MERGE (ds)-[:CONFIG_STATUS]->(cs)
MERGE (vt:Vdatastoretype {name:row.Type})
MERGE (ds)-[:DATASTORE_TYPE]->(vt)
WITH row,ds,vc,split(row.Hosts,',') as hosts
UNWIND hosts as vmhost
MATCH (vmh:Vspherehost {name:vmhost,managedby:row.`VI SDK UUID`})
MERGE (vmh)-[:CONNECTED_DATASTORE]->(ds);

// SECTION  MERGE (:Vdatastore) from vDatastore tab
CALL apoc.load.xls("path-to-vmware-import-file",'vDisk',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vd:Virtualdisk {path:row.Path})
SET vd.disk=row.Disk,vd.capacity=row.`Capacity MB`,vd.thin=row.Thin,vd.controller=row.Controller
SET vd.mode=row.`Disk Mode`,vd.eager=row.`Eagerly Scrub`,vd.template=row.Template
MERGE (vd)-[:VDISK_FOR_VM]-(vm)
with row,vc,vd,replace(split(row.Path,']')[0],'[','') as dsname
MATCH (ds:Vdatastore {name:dsname, managedby:row.`VI SDK UUID`})--(:Vspherehost {name:row.Host,managedby:row.`VI SDK UUID`})
MERGE (vd)-[:ON_DATASTORE]-(ds);

// SECTION  MERGE (:Vmadapter) from vNetwork
CALL apoc.load.xls("path-to-vmware-import-file",'vNetwork',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vmn:Vmadapter {mac:row.`Mac Address`,vmuuid:row.`VM UUID`})
MERGE (vmn)-[:ADAPTER_FOR]-(vm)
MERGE (vmat:Vmadaptertype {name:row.Adapter})
MERGE (vmn)-[:ADAPTER_TYPE]-(vmat)
SET vmn.startconnected=row.`Starts Connected`,vmn.ip=row.`IP Address`
with row,vmn
MATCH (pg:Vhostportgroup {name:row.Network,host:row.Host,managedby:row.`VI SDK UUID`})
MERGE (vmn)-[:IN_PORTGROUP]->(pg);

// SECTION  MERGE (:Vpartition) from vPartition
CALL apoc.load.xls("path-to-vmware-import-file",'vPartition',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vmp:Vpartition {disk:row.Disk,vmuuid:row.`VM UUID`})
MERGE (vmp)-[:PARTITION_FOR]-(vm)
SET vmp.capacity=row.`Capacity MB`,vmp.consumed=row.`Consumed MB`,vmp.free=row.`Free %`;

// SECTION  MERGE (:Vpartition) from vPartition
CALL apoc.load.xls("path-to-vmware-import-file",'vSnapshot',{header:true}) yield map as row
MATCH (vc:Vcenterserver {name:row.`VI SDK Server`}),(vm:Virtualmachine {uuid:row.`VM UUID`,managedby:row.`VI SDK UUID`})
MERGE (vmss:Vsnapshot {name:row.Name,vmuuid:row.`VM UUID`})
MERGE (vmss)-[:SNAPSHOT_OF]-(vm)
SET vmss.description=row.Description,vmss.timestamp=row.`Date / time`,vmss.size=row.`Size MB (total)`;

// SECTION Identify .unverified property for any existing top-level nodes managed by this vCenter
// Since the unverified property wasn't removed (above) this is an orphaned node and should be removed
CALL apoc.load.xls("path-to-vmware-import-file",'vCluster',{header:true}) yield map as row
MATCH (vc:Vcenterserver {uid:row.`VI SDK UUID`}) 
MATCH (n) where n.managedby=vc.uid and n.unverified=true
DETACH DELETE n;