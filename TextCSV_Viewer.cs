using System;
using System.IO;
using System.Windows.Forms;

namespace FileProcessing
{
    public partial class frmTextView : Form
    {
        public frmTextView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// ปุ่มโหลดไฟล์ข้อความแบบธรรมดา (Text View)
        /// </summary>
        private void btRead_Click(object sender, EventArgs e)
        {
            if (!File.Exists(tbFileName.Text)) return;
            string content = File.ReadAllText(tbFileName.Text);
            rtbShow.Text = content;
        }

        /// <summary>
        /// ปุ่มโหลดและจัดเรียงไฟล์ CSV พร้อมตัวกรองอัจฉริยะ (CSV View)
        /// </summary>
        private void btReadCSV_Click(object sender, EventArgs e)
        {
            // 1. ตรวจสอบไฟล์เพื่อป้องกันโปรแกรมแครช
            if (!File.Exists(tbFileName.Text))
            {
                MessageBox.Show("ไม่พบไฟล์ที่ระบุ กรุณาตรวจสอบพาธของไฟล์อีกครั้ง", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. ล้างข้อมูลเก่าและ "ล้างคอลัมน์เก่า" ทั้งหมด เพื่อให้ระบบสร้างคอลัมน์ใหม่ตามโครงสร้างไฟล์จริง
            dgvData.Rows.Clear();
            dgvData.Columns.Clear();

            // 3. อ่านค่าจากช่วงบรรทัด M และ N (ถ้าว่างให้ใส่ค่าเริ่มต้นอัตโนมัติ)
            int m = string.IsNullOrWhiteSpace(textBox1.Text) ? 1 : int.Parse(textBox1.Text);
            int n = string.IsNullOrWhiteSpace(textBox2.Text) ? int.MaxValue : int.Parse(textBox2.Text);

            // 4. ดึงค่าจาก ComboBox อัตโนมัติ (ไม่ว่าจะซ่อนอยู่ในแท็บไหนก็ตาม)
            string filterType = "";
            Action<Control> findComboBox = null;
            findComboBox = (parent) =>
            {
                foreach (Control c in parent.Controls)
                {
                    if (c is ComboBox cb)
                    {
                        // ตัดเครื่องหมายคำพูด และดักตัดเครื่องหมายจุด เช่น .dll หรือ .exe ให้เหลือแค่ชื่อคลีนๆ เสมอ
                        filterType = cb.Text.Trim().ToLower().Replace("\"", "").Replace(".", "");
                        return;
                    }
                    if (c.HasChildren) findComboBox(c);
                }
            };
            findComboBox(this);

            // ดัก Error กรณีป้อนช่วงเงื่อนไขผิดพลาด
            if (n < m)
            {
                MessageBox.Show("เงื่อนไขไม่ถูกต้อง: ค่า N ห้ามต่ำกว่าค่า M", "Invalid Range", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 5. ใช้ StreamReader อ่านไฟล์ทีละบรรทัดเพื่อความรวดเร็วและประหยัดแรม
                using (StreamReader srReader = new StreamReader(tbFileName.Text))
                {
                    string strLine;
                    bool bHeaderRead = false;
                    int currentLineIndex = 0;

                    while ((strLine = srReader.ReadLine()) != null)
                    {
                        // ข้ามบรรทัดคอมเมนต์ของระบบ MalwareBazaar ที่ขึ้นต้นด้วยเครื่องหมาย #
                        if (strLine.StartsWith("#")) continue;
                        if (string.IsNullOrWhiteSpace(strLine)) continue;

                        // แยกชิ้นส่วนข้อมูลด้วยเครื่องหมาย Comma (,) พร้อมตัดช่องว่างและเครื่องหมายคำพูด "" ออกให้หมดจด
                        string[] strValues_arr = strLine.Split(',');
                        for (int i = 0; i < strValues_arr.Length; i++)
                        {
                            strValues_arr[i] = strValues_arr[i].Trim().Trim('"');
                        }

                        // 🔥 ไฮไลท์แก้บั๊ก: สร้างหัวตาราง (Header) อัตโนมัติ ครบทุกคอลัมน์ที่มีอยู่ในไฟล์จริง!
                        if (!bHeaderRead)
                        {
                            foreach (string strHeader in strValues_arr)
                            {
                                // แอดหัวคอลัมน์เข้าตารางตรงๆ ตามที่ดึงได้จากแถวแรกของไฟล์
                                dgvData.Columns.Add(strHeader, strHeader);
                            }
                            bHeaderRead = true;
                            continue;
                        }

                        // นับบรรทัดข้อมูลจริง
                        currentLineIndex++;

                        // ตรวจสอบเงื่อนไข Partial Loading (M ถึง N)
                        if (currentLineIndex < m) continue;
                        if (currentLineIndex > n) break;

                        // 6. ระบบคัดกรองประเภทไฟล์แบบยืดหยุ่นสูง (Case-Insensitive ค้นหาเจอหมดทั้งตัวเล็กและตัวใหญ่)
                        if (!string.IsNullOrEmpty(filterType) && filterType != "all" && filterType != "all files")
                        {
                            bool isMatch = false;
                            foreach (string value in strValues_arr)
                            {
                                // คลีนข้อมูลในตารางมาเทียบกับตัวฟิลเตอร์ โดยตัดเครื่องหมายจุดออกด้วยเพื่อความแม่นยำ
                                string cleanValue = value.Trim().ToLower().Replace(".", "");
                                if (cleanValue.Equals(filterType, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatch = true;
                                    break;
                                }
                            }
                            if (!isMatch) continue; // หากแถวนี้ไม่มีประเภทไฟล์ที่เลือก ให้ข้ามแถวนี้ไปเลย
                        }

                        // 7. นำข้อมูลที่จัดเรียงช่องตรงกับจำนวนคอลัมน์แล้ว แอดลงตาราง DataGridView
                        if (strValues_arr.Length <= dgvData.Columns.Count)
                        {
                            dgvData.Rows.Add(strValues_arr);
                        }
                        else
                        {
                            // ป้องกันกรณีแถวข้อมูลยาวเกินคอลัมน์ ให้ตัดเอาเฉพาะความยาวที่เท่ากับคอลัมน์ตารางพอดี
                            string[] truncatedArray = new string[dgvData.Columns.Count];
                            Array.Copy(strValues_arr, truncatedArray, dgvData.Columns.Count);
                            dgvData.Rows.Add(truncatedArray);
                        }
                    }
                }

                // 8. ความสวยงามเพิ่มเติม: สั่งให้คอลัมน์ขยายขนาดกว้างอัตโนมัติตามความยาวเนื้อหาข้อมูลให้มองเห็นครบถ้วน
                foreach (DataGridViewColumn col in dgvData.Columns)
                {
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }

                // สรุปแจ้งเตือนเมื่อไม่พบผลลัพธ์
                if (dgvData.Rows.Count == 0)
                {
                    MessageBox.Show("ไม่พบข้อมูลที่ตรงกับเงื่อนไขการค้นหา", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"เกิดข้อผิดพลาดขณะโหลดไฟล์: {ex.Message}", "Crash Prevention", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    tbFileName.Text = ofd.FileName;
                }
            }
        }

        private void label3_Click(object sender, EventArgs e) { }
        private void label2_Click(object sender, EventArgs e) { }
        private void label4_Click(object sender, EventArgs e) { }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}