#Requires -Version 7

<#
    .SYNOPSIS
        Shared fixtures for the ConvertXmlToExcel tests.

    .DESCRIPTION
        Dot-source this file in a test's BeforeAll to get small XML builders
        that produce the exact shape the row builders and date readers consume.
        Keeping the fixtures here means the tests do not depend on the sample
        files in 'Tests\TestData' and can build the odd cases (a file spanning
        two months, a record with no date) on demand.
#>

function New-BatchXmlHC {
    <#
        .SYNOPSIS
            Build a minimal batch XML document with one delivery per date given.

        .PARAMETER LoadStartDate
            One or more 'load_start_date' values. One delivery is created per
            value. Pass an empty string to create a delivery without a date.

        .PARAMETER WithoutBatches
            Build every delivery with an empty 'batches' element.

            The month of a batch file is read from the delivery date, but the
            rows come from the batches under that delivery. A delivery without
            batches therefore gives a month that holds no rows at all, which is
            a valid file and not a failure.

        .EXAMPLE
            New-BatchXmlHC -LoadStartDate '2024-08-16T09:00:00', '2024-09-02T08:00:00'

        .EXAMPLE
            New-BatchXmlHC -LoadStartDate '2024-08-16T09:00:00' -WithoutBatches

            A file with a delivery loaded in August that produces no rows.
    #>
    param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [String[]]$LoadStartDate,
        [Switch]$WithoutBatches
    )

    $batches = if ($WithoutBatches) {
        '<batches></batches>'
    }
    else {
        @"
          <batches>
            <batch>
              <batchHeader>
                <batch_id>B1</batch_id>
                <batch_id_nr>1</batch_id_nr>
                <qty>10</qty>
              </batchHeader>
              <batchItems>
                <batchItem>
                  <material_code>0012</material_code>
                  <material_name>Sand</material_name>
                  <material_amount>500</material_amount>
                  <dosingOperationTimes>
                    <dosing>
                      <material_dosing_start_time>2024-08-16T09:31:00</material_dosing_start_time>
                      <material_dosing_end_time>2024-08-16T09:31:20</material_dosing_end_time>
                    </dosing>
                  </dosingOperationTimes>
                </batchItem>
              </batchItems>
            </batch>
          </batches>
"@
    }

    $deliveries = foreach ($date in $LoadStartDate) {
        @"
        <delivery>
          <deliveryHeader>
            <load_id_erp>ERP-1</load_id_erp>
            <!-- An order number with leading zeros and an identifier too long
                 to survive a double: both are damaged the moment they are
                 turned into a number, so every batch file built here carries
                 the case. -->
            <load_order_number>0000123456</load_order_number>
            <load_id_bcc>1234567890123456789</load_id_bcc>
            <load_mix_name>Mix</load_mix_name>
            <load_start_date>$date</load_start_date>
            <load_end_date></load_end_date>
            <load_qty>10</load_qty>
          </deliveryHeader>
          $batches
        </delivery>
"@
    }

    [xml]@"
<plant>
  <plantHeader>
    <country_code>BE</country_code>
    <company_code>1</company_code>
    <company_name>Contoso</company_name>
    <plant_code>P1</plant_code>
    <plant_name>Plant 1</plant_name>
  </plantHeader>
  <batchComputers>
    <batchComputer>
      <batchComputerHeader>
        <system_type>T</system_type>
        <system_provider>P</system_provider>
        <mixer_name>M</mixer_name>
        <offset>0</offset>
        <system_id>1</system_id>
        <mixer_size>2</mixer_size>
        <extraction_id>10</extraction_id>
        <file_created_on>2024-08-16T09:00:00</file_created_on>
      </batchComputerHeader>
      <deliveries>
        $deliveries
      </deliveries>
    </batchComputer>
  </batchComputers>
</plant>
"@
}

function New-AlarmXmlHC {
    <#
        .SYNOPSIS
            Build a minimal alarm XML document with one alarm per raised date.

        .PARAMETER Raised
            One or more 'raised' values, one alarm each. Empty string makes an
            alarm without a raised date.
    #>
    param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [String[]]$Raised
    )

    $alarms = foreach ($date in $Raised) {
        @"
        <alarm>
          <Id>1</Id>
          <text>Boom</text>
          <raised>$date</raised>
        </alarm>
"@
    }

    [xml]@"
<plant>
  <plantHeader>
    <country_code>BE</country_code>
    <company_code>1</company_code>
    <company_name>Contoso</company_name>
    <plant_code>P1</plant_code>
    <plant_name>Plant 1</plant_name>
  </plantHeader>
  <batchComputers>
    <batchComputer>
      <batchComputerHeader>
        <extraction_id>10</extraction_id>
        <file_created_on>2024-08-16T09:00:00</file_created_on>
      </batchComputerHeader>
      <alarms>
        $alarms
      </alarms>
    </batchComputer>
  </batchComputers>
</plant>
"@
}

function New-SequenceXmlHC {
    <#
        .SYNOPSIS
            Build a minimal sequence XML document with one batch computer per
            creation date given.

        .DESCRIPTION
            For a sequence file the month is decided per batch computer, on
            'file_created_on'. One batch computer is created per value, so a
            file spanning two months is made by passing two dates.

        .PARAMETER FileCreatedOn
            One or more 'file_created_on' values, one batch computer each. Pass
            an empty string to create a batch computer without a date.

        .EXAMPLE
            New-SequenceXmlHC -FileCreatedOn '2024-06-21T03:55:15+02:00'
    #>
    param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [String[]]$FileCreatedOn
    )

    $batchComputers = foreach ($date in $FileCreatedOn) {
        @"
    <batchComputer>
      <batchComputerHeader>
        <system_type>BCC</system_type>
        <system_provider>CACP</system_provider>
        <system_id>-</system_id>
        <mixer_name>VVM 1 Menger</mixer_name>
        <mixer_size>4.5</mixer_size>
        <extraction_id>10</extraction_id>
        <file_created_on>$date</file_created_on>
        <offset>+02:00</offset>
      </batchComputerHeader>
      <sequenceparameters>
        <sequenceparameter>
          <ID>3</ID>
          <name>Kalibratie 1</name>
          <basename>Kalibratie 1</basename>
          <NrofSubBatches>9</NrofSubBatches>
          <maxbatchsize>4.5</maxbatchsize>
          <maxbatchsizeunit>m3</maxbatchsizeunit>
          <blocked>FALSE</blocked>
          <parameters>
            <parameter>
              <properties>
                <property>
                  <name>MotorDischarge1</name>
                  <prefix>Toerental trap 1</prefix>
                  <value>25</value>
                  <itempath>\Plant\VVM 1 Menger\Engine</itempath>
                  <uppath>\Menger VVM1\Lossen</uppath>
                  <suffix>%</suffix>
                </property>
              </properties>
            </parameter>
          </parameters>
          <subBatches>
            <subBatch>
              <name>SubBatch 1</name>
              <properties>
                <property>
                  <name>Water</name>
                  <prefix>Water</prefix>
                  <value>120</value>
                  <itempath>\Plant\VVM 1 Menger\Water</itempath>
                  <uppath>\Menger VVM1\Water</uppath>
                  <suffix>l</suffix>
                </property>
              </properties>
            </subBatch>
          </subBatches>
        </sequenceparameter>
      </sequenceparameters>
    </batchComputer>
"@
    }

    [xml]@"
<plant>
  <plantHeader>
    <country_code>NL</country_code>
    <company_code>NL30</company_code>
    <company_name>Mebin B.V.</company_name>
    <plant_code>NL3D</plant_code>
    <plant_name>Amsterdam Zuid</plant_name>
  </plantHeader>
  <batchComputers>
    $batchComputers
  </batchComputers>
</plant>
"@
}