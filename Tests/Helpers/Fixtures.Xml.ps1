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

        .EXAMPLE
            New-BatchXmlHC -LoadStartDate '2024-08-16T09:00:00', '2024-09-02T08:00:00'
    #>
    param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [String[]]$LoadStartDate
    )

    $deliveries = foreach ($date in $LoadStartDate) {
        @"
        <delivery>
          <deliveryHeader>
            <load_id_erp>ERP-1</load_id_erp>
            <load_mix_name>Mix</load_mix_name>
            <load_start_date>$date</load_start_date>
            <load_end_date></load_end_date>
            <load_qty>10</load_qty>
          </deliveryHeader>
          <batches>
            <batch>
              <batchHeader>
                <batch_id>B1</batch_id>
                <batch_id_nr>1</batch_id_nr>
                <qty>10</qty>
              </batchHeader>
              <batchItems></batchItems>
            </batch>
          </batches>
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