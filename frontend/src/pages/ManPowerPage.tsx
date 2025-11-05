import React, { useMemo, useRef, useState } from "react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogClose,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Trash2, Plus, Download, Upload, Eye } from "lucide-react"; // NEW: Eye icon
import * as XLSX from "xlsx";
import { AnimatePresence, motion } from "framer-motion";

const sections = [
  { section_code: "K2FM00", section_name: "FM2" },
  { section_code: "K2PE00", section_name: "PE" },
  { section_code: "K2MN00", section_name: "MTN" },
  { section_code: "K2QA00", section_name: "QA" },
  { section_code: "K2SC00", section_name: "SCM" },
  { section_code: "K2TL00", section_name: "TL" },
  { section_code: "K2PD06", section_name: "PD6" },
  { section_code: "K2QC00", section_name: "QC" },
  { section_code: "K1AC00", section_name: "ACC" },
  { section_code: "K1CM00", section_name: "CM" },
  { section_code: "K1IT00", section_name: "IT" },
  { section_code: "K1SS00", section_name: "SYS" },
  { section_code: "K1ED00", section_name: "EEE" },
  { section_code: "K1EE00", section_name: "EEE" },
  { section_code: "K1EG00", section_name: "EEE" },
  { section_code: "K1EM00", section_name: "EEE" },
  { section_code: "K1EX00", section_name: "EXE" },
  { section_code: "K1CG00", section_name: "CG" },
  { section_code: "K1QS00", section_name: "QMS" },
  { section_code: "K1ND00", section_name: "ND" },
  { section_code: "K1OV00", section_name: "OVT" },
  { section_code: "K1RD00", section_name: "RD" },
  { section_code: "K1SH00", section_name: "SHE" },
  { section_code: "K1SM00", section_name: "SM" },
  { section_code: "K1PJ00", section_name: "PJC" },
  { section_code: "K1PU00", section_name: "PUR" },
  { section_code: "K1SQ00", section_name: "SQE" },
  { section_code: "K1LG00", section_name: "SCM" },
  { section_code: "K1PS00", section_name: "SCM" },
  { section_code: "K1SC00", section_name: "SCM" },
  { section_code: "K1CE00", section_name: "CE" },
  { section_code: "K1PE00", section_name: "PE" },
  { section_code: "K1TL00", section_name: "TL" },
  { section_code: "K1QA00", section_name: "QA" },
  { section_code: "K1QI00", section_name: "QI" },
  { section_code: "K1TP00", section_name: "TPS" },
  { section_code: "K1MN00", section_name: "MTN" },
  { section_code: "K1UT00", section_name: "UT" },
  { section_code: "K1QC00", section_name: "QC" },
  { section_code: "K1QG00", section_name: "QG" },
  { section_code: "K1MA01", section_name: "MA1" },
  { section_code: "K1A000", section_name: "PD1" },
  { section_code: "K1B000", section_name: "PD1B" },
  { section_code: "K1M000", section_name: "PD2" },
  { section_code: "K1G000", section_name: "PD5" },
  { section_code: "K1MA02", section_name: "MA2" },
  { section_code: "K1GC00", section_name: "GDC" },
  { section_code: "K1CH00", section_name: "PD3C" },
  { section_code: "K1MC00", section_name: "PD3" },
  { section_code: "K1PD04", section_name: "PD4" },
  { section_code: "K1MA03", section_name: "MA3" },
  { section_code: "K1PD07", section_name: "PD7" },
  { section_code: "K1C000", section_name: "PD7" },
  { section_code: "K1I000", section_name: "PD7" },
  { section_code: "K1T000", section_name: "PD7" },
];

type EmployeeType = "admin" | "direct" | "indirect";

interface ManpowerData {
  emp_type: EmployeeType;
  permanent: number;
  subcontract: number;
  student: number;
}

interface SectionData {
  id: string;
  section_code: string;
  items: ManpowerData[];
}

// NEW: types for preview
type Block = { perm: number; sub: number; stu: number; total: number };
type PreviewRow = {
  section_code: string;
  section_name: string;
  admin: Block;
  direct: Block;
  indirect: Block;
  grandTotal: number;
};

function ManPowerPage() {
  const [sectionList, setSectionList] = useState<SectionData[]>([
    {
      id: "1",
      section_code: "",
      items: [{ emp_type: "admin", permanent: 0, subcontract: 0, student: 0 }],
    },
  ]);

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [step, setStep] = useState<"editor" | "preview">("editor");

  const addSection = () => {
    const newSection: SectionData = {
      id: Date.now().toString(),
      section_code: "",
      items: [{ emp_type: "admin", permanent: 0, subcontract: 0, student: 0 }],
    };
    setSectionList((prev) => [...prev, newSection]);
  };

  const removeSection = (id: string) => {
    setSectionList((prev) => prev.filter((section) => section.id !== id));
  };

  const updateSectionCode = (id: string, code: string) => {
    setSectionList((prev) =>
      prev.map((section) =>
        section.id === id ? { ...section, section_code: code } : section
      )
    );
  };

  const toggleEmployeeType = (sectionId: string, empType: EmployeeType) => {
    setSectionList((prev) =>
      prev.map((section) => {
        if (section.id === sectionId) {
          const exists = section.items.find(
            (item) => item.emp_type === empType
          );
          if (exists) {
            if (section.items.length > 1) {
              return {
                ...section,
                items: section.items.filter(
                  (item) => item.emp_type !== empType
                ),
              };
            }
          } else {
            return {
              ...section,
              items: [
                ...section.items,
                { emp_type: empType, permanent: 0, subcontract: 0, student: 0 },
              ],
            };
          }
        }
        return section;
      })
    );
  };

  const updateManpower = (
    sectionId: string,
    empType: EmployeeType,
    field: "permanent" | "subcontract" | "student",
    value: number
  ) => {
    setSectionList((prev) =>
      prev.map((section) => {
        if (section.id === sectionId) {
          return {
            ...section,
            items: section.items.map((item) =>
              item.emp_type === empType ? { ...item, [field]: value } : item
            ),
          };
        }
        return section;
      })
    );
  };

  const handleSubmit = () => {
    console.log("Submitted data:", sectionList);
    alert("Data saved successfully!");
  };

  // ===================== Upload Excel =====================
  const COL = { O: 14, Q: 16, S: 18 };

  const normalizeEmpType = (val: any): EmployeeType | null => {
    if (!val) return null;
    const t = String(val).trim().toLowerCase();
    if (["admin", "adm", "office"].includes(t)) return "admin";
    if (["direct", "dir"].includes(t)) return "direct";
    if (["indirect", "indir", "in-direct", "ind"].includes(t))
      return "indirect";
    return null;
  };

  const mapContractToField = (
    val: any
  ): "permanent" | "subcontract" | "student" | null => {
    if (!val) return null;
    const t = String(val).trim().toUpperCase();
    if (t === "PERM") return "permanent";
    if (t === "SUB") return "subcontract";
    if (t === "STU") return "student";
    return null;
  };

  const buildSectionListFromAgg = (
    agg: Record<string, Record<EmployeeType, ManpowerData>>
  ) => {
    const list: SectionData[] = [];
    Object.entries(agg).forEach(([section_code, byType]) => {
      const items = Object.values(byType).filter(
        (item) => item.permanent > 0 || item.subcontract > 0 || item.student > 0
      );
      if (items.length > 0) {
        list.push({
          id: `${section_code}-${Date.now()}-${Math.random()}`,
          section_code,
          items,
        });
      }
    });
    return list;
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    try {
      const file = e.target.files?.[0];
      if (!file) return;

      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });

      const wsName = wb.SheetNames[0];
      const ws = wb.Sheets[wsName];

      const rows: any[][] = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        raw: true,
        defval: "",
      });

      if (rows.length <= 1) {
        alert("ไฟล์ไม่มีข้อมูล (ต้องมีอย่างน้อย 2 แถวขึ้นไป)");
        return;
      }

      const agg: Record<string, Record<EmployeeType, ManpowerData>> = {};

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const section_code = String(row[COL.O] || "").trim();
        const empTypeRaw = row[COL.Q];
        const contractRaw = row[COL.S];

        if (!section_code) continue;

        const emp_type = normalizeEmpType(empTypeRaw);
        const field = mapContractToField(contractRaw);

        if (!emp_type || !field) continue;

        if (!agg[section_code]) {
          agg[section_code] = {
            admin: {
              emp_type: "admin",
              permanent: 0,
              subcontract: 0,
              student: 0,
            },
            direct: {
              emp_type: "direct",
              permanent: 0,
              subcontract: 0,
              student: 0,
            },
            indirect: {
              emp_type: "indirect",
              permanent: 0,
              subcontract: 0,
              student: 0,
            },
          };
        }

        agg[section_code][emp_type][field] += 1;
      }

      const list = buildSectionListFromAgg(agg);
      if (list.length === 0) {
        alert("ไม่พบข้อมูลที่แมปได้จากคอลัมน์ O/Q/S");
        return;
      }

      setSectionList(list);
      alert("อัปโหลดและแปลงข้อมูลสำเร็จ!");
    } catch (err) {
      console.error(err);
      alert("เกิดข้อผิดพลาดในการอ่านไฟล์ Excel");
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };
  // ===================== End Upload Excel =====================

  const handleDownloadTemplate = () => {
    const link = document.createElement("a");
    link.href = "/assets/manpower_template.xlsx";
    link.download = "manpower_template.xlsx";
    link.click();
  };

  // =============== Preview helpers (NEW) ===============
  const sectionMap = useMemo(
    () =>
      new Map(
        sections.map((s) => [s.section_code.trim(), s.section_name.trim()])
      ),
    []
  );

  const emptyBlock = (): Block => ({ perm: 0, sub: 0, stu: 0, total: 0 });

  const buildPreviewRows = (data: SectionData[]): PreviewRow[] => {
    const map = new Map<string, PreviewRow>();

    for (const sec of data) {
      const code = sec.section_code?.trim() || "";
      if (!code) continue;

      if (!map.has(code)) {
        map.set(code, {
          section_code: code,
          section_name: sectionMap.get(code) || "-",
          admin: emptyBlock(),
          direct: emptyBlock(),
          indirect: emptyBlock(),
          grandTotal: 0,
        });
      }
      const row = map.get(code)!;

      for (const item of sec.items) {
        const target = row[item.emp_type];
        target.perm += item.permanent || 0;
        target.sub += item.subcontract || 0;
        target.stu += item.student || 0;
        target.total = target.perm + target.sub + target.stu;
      }

      row.grandTotal = row.admin.total + row.direct.total + row.indirect.total;
    }

    // เรียงตาม section_code
    return Array.from(map.values()).sort((a, b) =>
      a.section_code.localeCompare(b.section_code)
    );
  };

  const previewRows = useMemo(
    () => buildPreviewRows(sectionList),
    [sectionList]
  );

  const totalRow = useMemo(() => {
    const t = {
      admin: emptyBlock(),
      direct: emptyBlock(),
      indirect: emptyBlock(),
      grandTotal: 0,
    };
    previewRows.forEach((r) => {
      t.admin.perm += r.admin.perm;
      t.admin.sub += r.admin.sub;
      t.admin.stu += r.admin.stu;
      t.admin.total += r.admin.total;

      t.direct.perm += r.direct.perm;
      t.direct.sub += r.direct.sub;
      t.direct.stu += r.direct.stu;
      t.direct.total += r.direct.total;

      t.indirect.perm += r.indirect.perm;
      t.indirect.sub += r.indirect.sub;
      t.indirect.stu += r.indirect.stu;
      t.indirect.total += r.indirect.total;

      t.grandTotal += r.grandTotal;
    });
    return t;
  }, [previewRows]);
  // =============== End Preview helpers (NEW) ===============

  return (
    <div className="">
      <Dialog>
        <DialogTrigger asChild>
          <Button variant="outline">Open Dialog</Button>
        </DialogTrigger>
        <DialogContent
          className="min-w-[950px] w-full max-h-[90vh] overflow-y-auto  !bg-gray-900 text-white border-gray-700"
          style={{ backgroundColor: "#111827" }}
        >
          <DialogHeader>
            <DialogTitle className="text-2xl text-white">
              Add Allocation Manpower
            </DialogTitle>
            <DialogDescription className="text-gray-400">
              Make changes to your manpower here. Click save when you're done.
            </DialogDescription>
          </DialogHeader>

          <div className="flex flex-wrap gap-3 justify-end mb-4">
            <Button
              variant="outline"
              className="gap-2 border-gray-600 text-white hover:bg-gray-800"
              onClick={handleDownloadTemplate}
            >
              <Download className="w-4 h-4" />
              Download Template
            </Button>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={handleFileChange}
            />
            <Button
              className="bg-green-600 hover:bg-green-700 text-white gap-2"
              onClick={() => fileInputRef.current?.click()}
            >
              <Upload className="w-4 h-4" />
              Upload Excel
            </Button>

            <div className="flex items-center gap-2 mb-3">
              <button
                className={`px-3 py-1 rounded-full text-sm ${
                  step === "editor"
                    ? "bg-blue-600 text-white"
                    : "bg-gray-800 text-gray-300"
                }`}
                onClick={() => setStep("editor")}
              >
                Editor
              </button>
              <span className="text-gray-500">→</span>
              <button
                className={`px-3 py-1 rounded-full text-sm ${
                  step === "preview"
                    ? "bg-blue-600 text-white"
                    : "bg-gray-800 text-gray-300"
                }`}
                onClick={() => setStep("preview")}
              >
                Preview
              </button>
            </div>
          </div>

          {/* พื้นที่สลับหน้าแบบ animate */}
          <div className="relative">
            <AnimatePresence mode="wait">
              {step === "editor" ? (
                <motion.div
                  key="editor"
                  initial={{ opacity: 0, y: 8 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -8 }}
                  transition={{ duration: 0.18, ease: "easeOut" }}
                >
                  {/* === Existing editor UI === */}
                  <div className="space-y-6">
                    {sectionList.map((section, sectionIndex) => (
                      <div
                        key={section.id}
                        className="border-2 border-gray-700 rounded-lg p-6 relative "
                      >
                        <div className="absolute -top-4 left-4 bg-blue-600 text-white px-6 py-2 rounded-full font-semibold text-base">
                          Section #{sectionIndex + 1}
                        </div>

                        {sectionList.length > 1 && (
                          <button
                            type="button"
                            onClick={() => removeSection(section.id)}
                            className="absolute -top-3 -right-3 bg-red-500 hover:bg-red-600 text-white rounded-full p-3 shadow-lg transition-colors"
                          >
                            <Trash2 className="w-5 h-5" />
                          </button>
                        )}

                        <div className="mt-4 mb-6">
                          <Label className="text-lg font-semibold mb-2 block text-white">
                            Section Code
                          </Label>
                          <Select
                            value={section.section_code}
                            onValueChange={(value) =>
                              updateSectionCode(section.id, value)
                            }
                          >
                            <SelectTrigger
                              className="w-full h-12 text-base !bg-gray-900 border-gray-600 text-white"
                              style={{ backgroundColor: "#111827" }}
                            >
                              <SelectValue placeholder="Select section code" />
                            </SelectTrigger>
                            <SelectContent
                              className="!bg-gray-800 border-gray-700"
                              style={{ backgroundColor: "#1f2937" }}
                            >
                              {sections.map((sec) => (
                                <SelectItem
                                  key={sec.section_code}
                                  value={sec.section_code}
                                  className="text-white hover:!bg-gray-700 focus:!bg-gray-700"
                                  style={{ backgroundColor: "transparent" }}
                                >
                                  {sec.section_code} - {sec.section_name}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>

                        <div className="mb-6">
                          <Label className="text-lg font-semibold mb-3 block text-white">
                            Employee Type
                          </Label>
                          <div className="flex gap-2">
                            {["admin", "direct", "indirect"].map((type) => {
                              const hasType = section.items.some(
                                (item) => item.emp_type === type
                              );
                              return (
                                <Button
                                  key={type}
                                  type="button"
                                  variant="outline"
                                  className={`px-6 py-2 ${
                                    hasType
                                      ? "bg-blue-600 hover:bg-blue-700 text-white border-blue-600"
                                      : "bg-transparent border-gray-600 text-white hover:bg-gray-800"
                                  }`}
                                  onClick={() =>
                                    toggleEmployeeType(
                                      section.id,
                                      type as EmployeeType
                                    )
                                  }
                                >
                                  {type.charAt(0).toUpperCase() + type.slice(1)}
                                </Button>
                              );
                            })}
                          </div>
                        </div>

                        <div className="space-y-4 bg-white rounded-lg p-6">
                          {section.items.map((item) => (
                            <div
                              key={item.emp_type}
                              className="border-b border-gray-200 last:border-0 pb-4 last:pb-0"
                            >
                              <div className="grid grid-cols-4 gap-6 items-end">
                                <div>
                                  <Label className="text-sm font-medium mb-2 block text-gray-700">
                                    Employee Type
                                  </Label>
                                  <div className="text-xl font-bold text-gray-900">
                                    {item.emp_type.charAt(0).toUpperCase() +
                                      item.emp_type.slice(1)}
                                  </div>
                                </div>

                                <div>
                                  <Label className="text-sm font-medium mb-2 block text-gray-700">
                                    Permanent
                                  </Label>
                                  <Input
                                    type="number"
                                    min="0"
                                    value={item.permanent}
                                    onChange={(e) =>
                                      updateManpower(
                                        section.id,
                                        item.emp_type,
                                        "permanent",
                                        parseInt(e.target.value) || 0
                                      )
                                    }
                                    className="h-11 bg-white border-gray-300 text-gray-900"
                                  />
                                </div>

                                <div>
                                  <Label className="text-sm font-medium mb-2 block text-gray-700">
                                    Subcontract
                                  </Label>
                                  <Input
                                    type="number"
                                    min="0"
                                    value={item.subcontract}
                                    onChange={(e) =>
                                      updateManpower(
                                        section.id,
                                        item.emp_type,
                                        "subcontract",
                                        parseInt(e.target.value) || 0
                                      )
                                    }
                                    className="h-11 bg-white border-gray-300 text-gray-900"
                                  />
                                </div>

                                <div>
                                  <Label className="text-sm font-medium mb-2 block text-gray-700">
                                    Student
                                  </Label>
                                  <Input
                                    type="number"
                                    min="0"
                                    value={item.student}
                                    onChange={(e) =>
                                      updateManpower(
                                        section.id,
                                        item.emp_type,
                                        "student",
                                        parseInt(e.target.value) || 0
                                      )
                                    }
                                    className="h-11 bg-white border-gray-300 text-gray-900"
                                  />
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}

                    <Button
                      type="button"
                      onClick={addSection}
                      className="bg-blue-600 hover:bg-blue-700 text-white gap-2 px-6"
                    >
                      <Plus className="w-4 h-4" />
                      Add Section
                    </Button>
                  </div>
                </motion.div>
              ) : (
                <motion.div
                  key="preview"
                  initial={{ opacity: 0, y: 8 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -8 }}
                  transition={{ duration: 0.18, ease: "easeOut" }}
                >
                  <div className="border border-gray-700 rounded-lg">
                    <div className="max-h-[65vh] overflow-auto">
                      <table className=" w-full text-sm">
                        <thead className="bg-gray-800 sticky top-0 z-10">
                          <tr>
                            <th className="px-3 py-2 text-left">
                              section_code
                            </th>
                            <th className="px-3 py-2 text-left">
                              section name
                            </th>
                            <th className="px-3 py-2 text-center" colSpan={4}>
                              admin
                            </th>
                            <th className="px-3 py-2 text-center" colSpan={4}>
                              direct
                            </th>
                            <th className="px-3 py-2 text-center" colSpan={4}>
                              indirect
                            </th>
                            <th className="px-3 py-2 text-center">
                              Grand total
                            </th>
                          </tr>
                          <tr className="bg-gray-800">
                            <th className="px-3 py-2"></th>
                            <th className="px-3 py-2"></th>
                            <th className="px-3 py-2 text-center">perm</th>
                            <th className="px-3 py-2 text-center">sub</th>
                            <th className="px-3 py-2 text-center">stu</th>
                            <th className="px-3 py-2 text-center">total</th>
                            <th className="px-3 py-2 text-center">perm</th>
                            <th className="px-3 py-2 text-center">sub</th>
                            <th className="px-3 py-2 text-center">stu</th>
                            <th className="px-3 py-2 text-center">total</th>
                            <th className="px-3 py-2 text-center">perm</th>
                            <th className="px-3 py-2 text-center">sub</th>
                            <th className="px-3 py-2 text-center">stu</th>
                            <th className="px-3 py-2 text-center">total</th>
                            <th className="px-3 py-2 text-center"></th>
                          </tr>
                        </thead>

                        <tbody className="divide-y divide-gray-700">
                          {previewRows.map((r) => (
                            <tr
                              key={r.section_code}
                              className="hover:bg-gray-800/40"
                            >
                              <td className="px-3 py-2 font-mono">
                                {r.section_code}
                              </td>
                              <td className="px-3 py-2">{r.section_name}</td>

                              <td className="px-3 py-2 text-center">
                                {r.admin.perm}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {r.admin.sub}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {r.admin.stu}
                              </td>
                              <td className="px-3 py-2 text-center font-semibold">
                                {r.admin.total}
                              </td>

                              <td className="px-3 py-2 text-center">
                                {r.direct.perm}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {r.direct.sub}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {r.direct.stu}
                              </td>
                              <td className="px-3 py-2 text-center font-semibold">
                                {r.direct.total}
                              </td>

                              <td className="px-3 py-2 text-center">
                                {r.indirect.perm}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {r.indirect.sub}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {r.indirect.stu}
                              </td>
                              <td className="px-3 py-2 text-center font-semibold">
                                {r.indirect.total}
                              </td>

                              <td className="px-3 py-2 text-center font-bold">
                                {r.grandTotal}
                              </td>
                            </tr>
                          ))}
                        </tbody>

                        <tfoot className="bg-gray-800 sticky bottom-0">
                          <tr>
                            <td className="px-3 py-2 font-semibold" colSpan={2}>
                              TOTAL
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.admin.perm}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.admin.sub}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.admin.stu}
                            </td>
                            <td className="px-3 py-2 text-center font-bold">
                              {totalRow.admin.total}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.direct.perm}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.direct.sub}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.direct.stu}
                            </td>
                            <td className="px-3 py-2 text-center font-bold">
                              {totalRow.direct.total}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.indirect.perm}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.indirect.sub}
                            </td>
                            <td className="px-3 py-2 text-center font-semibold">
                              {totalRow.indirect.stu}
                            </td>
                            <td className="px-3 py-2 text-center font-bold">
                              {totalRow.indirect.total}
                            </td>
                            <td className="px-3 py-2 text-center font-extrabold">
                              {totalRow.grandTotal}
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          <DialogFooter className="mt-4 gap-2">
            {step === "preview" ? (
              <Button variant="outline" onClick={() => setStep("editor")}>
                Back to Editor
              </Button>
            ) : (
              <Button variant="outline" onClick={() => setStep("preview")}>
                Preview
              </Button>
            )}
            <DialogClose asChild>
              <Button
                variant="outline"
                className="border-gray-600 text-white hover:bg-gray-800"
              >
                Close
              </Button>
            </DialogClose>
            <Button
              onClick={handleSubmit}
              className="bg-blue-600 hover:bg-blue-700 text-white"
            >
              Save
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

export default ManPowerPage;
