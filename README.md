# Alpha Live Data

### Capital Requirements
<ul>
    <li>OPC Server</li>
    <ul>
        <li>Utilized RsLinx OPC Server</li>
        <ul>
            <li>PC and DDE clients are supported for any number of devices. It also supports applications developed for the RSLinx Classic C API. But note that this is limited to 32bit client only. Additional information is found within Rockwell's <a href="https://literature.rockwellautomation.com/idc/groups/literature/documents/gr/linx-gr001_-en-e.pdf"> documentation</a>.</li>
        </ul>
        <li>Other OPC servers may be utilized, but have not been tested.</li>
        <ul>
            <li>Matrikon OPC</li>
            <li>Cyberlogic</li>
            <li>Graybox</li>
        </ul>
    </ul>
    <li>Computer System configure with Python 3 or later</li>
    <li>Database</li>
    <ul>
    <li>SQL Express</li>
    <li>Oracle</li>
    <li>Microsoft Access</li>
    </ul>
</ul>

### Implementation
Identify the OPC server's IP address, all OPC tags should be hosted on the OPC server. Configure tags within the scrapping script as demonstrated. 