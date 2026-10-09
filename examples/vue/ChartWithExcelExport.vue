<!-- Vue 3 + highcharts-vue (install both in your app; they are not dependencies of this library). -->
<script setup lang="ts">
import { ref } from 'vue';
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { Chart } from 'highcharts-vue';
import { downloadHighchartsAsXlsx, installHighchartsExcelExport } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts);

const chartRef = ref<{ chart: Highcharts.Chart } | null>(null);
const options: Highcharts.Options = {
  title: { text: 'Daily signups' },
  xAxis: { categories: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'] },
  series: [{ type: 'line', name: 'Signups', data: [31, 42, 38, 51, 47] }],
};

async function download(): Promise<void> {
  if (chartRef.value) await downloadHighchartsAsXlsx(chartRef.value.chart, { filename: 'signups' });
}
</script>

<template>
  <Chart ref="chartRef" :highcharts="Highcharts" :options="options" />
  <button type="button" @click="download">Download editable Excel chart</button>
</template>
