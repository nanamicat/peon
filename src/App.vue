<template>
  <div style="background: #ececec; padding: 30px">
    <a-card :bordered="false" title="药娘苦力模拟器">
      <div class="flex">
        <div class="left">
          <div class="top">
            <a-upload
              :before-upload="selectWord"
              :show-upload-list="false"
              list-type="picture-card"
              name="avatar"
            >
              <div>
                <a-icon :type="word ? 'check' : 'file-word'" />
                <div class="ant-upload-text">Word 模板</div>
              </div>
            </a-upload>

            <a-upload
              :before-upload="selectExcel"
              :show-upload-list="false"
              list-type="picture-card"
              name="avatar"
            >
              <div>
                <a-icon :type="excel ? 'check' : 'bar-chart'" />
                <div class="ant-upload-text">Excel 数据表</div>
              </div>
            </a-upload>
          </div>

          <a-button
            v-if="progress === null"
            :disabled="!(word && excel)"
            icon="download"
            size="large"
            style="width: 100%"
            type="primary"
            @click="handleUpload()"
          >
            Go!
          </a-button>
          <a-progress v-else :percent="Math.round(progress * 100)" />
        </div>
        <vue-qrcode
          :options="{ width: 152, margin: 0 }"
          value="https://qr.alipay.com/fkx16235k24tcg2zfqluo16"
        ></vue-qrcode>
      </div>
    </a-card>
  </div>
</template>

<script lang="ts">
import {Component, Vue} from "vue-property-decorator";
import HelloWorld from "./components/HelloWorld.vue";
// @ts-ignore
import VueQrcode from "@chenfengyuan/vue-qrcode";

import XLSX from "xlsx";
import createReport from "docx-templates";
import JSZip from "jszip";
import {saveAs} from "file-saver";

@Component({
  components: {
    HelloWorld,
    VueQrcode,
  },
})
export default class App extends Vue {
  word: File | null = null;
  excel: File | null = null;
  progress: number | null = null;

  async handleUpload() {
    if (!this.word || !this.excel) return;
    if (this.progress !== null) return;
    this.progress = 0;

    const workbook = XLSX.read(await this.excel.arrayBuffer(), {
      type: "array",
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    let data: any[] = XLSX.utils.sheet_to_json(sheet, {
      range: "A4:ZZ9999",
      defval: "",
    });
    console.log(data);
    data = data.filter((d) =>
      [
        "轻度1",
        "轻度2",
        "轻度3",
        "轻度4",
        "中度1",
        "中度2",
        "中度3",
        "中度4",
      ].includes(d.模版)
    );

    console.log(data);

    const template = await this.word.arrayBuffer();

    const zip = new JSZip();

    for (const [index, row] of data.entries()) {
      this.progress = index / data.length;
      const report = await createReport({
        // @ts-ignore
        template,
        cmdDelimiter: ["{", "}"],
        data: row,
      });

      zip.file(`心理评测报告_${row.姓名}.docx`, report);
    }

    this.progress = 1;
    const content = await zip.generateAsync({ type: "blob" });

    saveAs(content, "report.zip");

    this.progress = null;
  }

  selectWord(file: File) {
    this.word = file;
    return false;
  }

  selectExcel(file: File) {
    this.excel = file;
    return false;
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}

.flex {
  display: flex;
}

.left {
  border-right: 1px black solid;
  margin-right: 16px;
  padding-right: 16px;
}

.top {
  display: flex;
  flex-direction: row;
}
</style>
