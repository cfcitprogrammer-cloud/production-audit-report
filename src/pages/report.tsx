"use client"

import * as React from "react"
import { format, addDays, isBefore, isSameDay, startOfDay } from "date-fns"
import {
  Calendar as CalendarIcon,
  FileSpreadsheet,
  Download,
  Loader2,
  Search,
} from "lucide-react"
import { supabase } from "@/lib/supabase"
import * as XLSX from "xlsx"

// shadcn/ui components
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import { Checkbox } from "@/components/ui/checkbox"
import { Label } from "@/components/ui/label"
import { Button } from "@/components/ui/button"
import { Calendar } from "@/components/ui/calendar"
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover"
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table"

// --- Type Definitions ---
interface SkuMergedRow {
  item_code: string
  item_description: string
  combined_value: number
  uom: string
}

export default function SkuFlatReportWithUomPage() {
  const [department, setDepartment] = React.useState<string>("")
  const [shifts, setShifts] = React.useState<{ day: boolean; night: boolean }>({
    day: true,
    night: true,
  })

  const [startDate, setStartDate] = React.useState<Date | undefined>(
    startOfDay(new Date())
  )
  const [endDate, setEndDate] = React.useState<Date | undefined>(
    startOfDay(new Date())
  )

  const [isLoading, setIsLoading] = React.useState<boolean>(false)
  const [mergedRows, setMergedRows] = React.useState<SkuMergedRow[] | null>(
    null
  )

  // --- XLSX Export Handler ---
  const handleExportToExcel = () => {
    if (!mergedRows || mergedRows.length === 0) return

    const worksheetData = mergedRows.map((row) => ({
      "SKU Code": row.item_code,
      "Item Description": row.item_description,
      "Total Extracted Value": row.combined_value,
      "Unit of Measure (UOM)": row.uom.toUpperCase(),
    }))

    const worksheet = XLSX.utils.json_to_sheet(worksheetData)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, "Yield Profile")

    worksheet["!cols"] = [{ wch: 18 }, { wch: 45 }, { wch: 22 }, { wch: 22 }]

    const startString = startDate ? format(startDate, "yyyyMMdd") : "start"
    const endString = endDate ? format(endDate, "yyyyMMdd") : "end"
    const fileName = `${department}_flat_report_${startString}_to_${endString}.xlsx`

    XLSX.writeFile(workbook, fileName)
  }

  const handleGenerateReport = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!startDate || !endDate || !department) return

    setIsLoading(true)
    setMergedRows(null)

    const generatedProdIds: string[] = []
    let currentIterDate = new Date(startDate)
    while (
      isBefore(currentIterDate, endDate) ||
      isSameDay(currentIterDate, endDate)
    ) {
      const dateString = format(currentIterDate, "yyyy-MM-dd")
      if (shifts.day) generatedProdIds.push(`PROD-${dateString}-day`)
      if (shifts.night) generatedProdIds.push(`PROD-${dateString}-night`)
      currentIterDate = addDays(currentIterDate, 1)
    }

    if (generatedProdIds.length === 0) {
      setMergedRows([])
      setIsLoading(false)
      return
    }

    try {
      const localAggregationMap = new Map<string, SkuMergedRow>()

      // Shared item specs lookup dictionary across all lines
      const globalDescMap = new Map<string, { desc: string; uom: string }>()

      const getOrCreateRow = (code: string): SkuMergedRow => {
        if (!localAggregationMap.has(code)) {
          const skuMeta = globalDescMap.get(code)
          localAggregationMap.set(code, {
            item_code: code,
            item_description: skuMeta?.desc || "Missing item specs",
            combined_value: 0,
            uom: skuMeta?.uom || "units",
          })
        }
        return localAggregationMap.get(code)!
      }

      const isAll = department === "all"

      // Execute queries conditioned on chosen division parameters
      const [
        // 1. Bihon Line Queries
        bhCooking,
        bhPacking,
        bhFg,
        bhSku,
        // 2. Snackfood Line Queries
        sfBlend,
        sfPremix,
        sfMix,
        sfFrying,
        sfFlavor,
        sfPiece,
        sfFg,
        sfSku,
        // 3. Catmon Line Queries
        cmMix,
        cmDry,
        cmPacking,
        cmFg,
        cmSku,
        // 4. Kingsforth Sotanghon / Shared Core Queries
        kfPacking,
        kfFg,
        kfSeasoning,
        kfSku,
        // 5. Kingsforth Specific Subdivision Queries
        kfHePacking,
        kfHeFg,
        kfCantonPacking,
        kfCantonFg,
        kfSfPacking,
        kfSfFg,
      ] = await Promise.all([
        // --- Bihon blocks ---
        isAll || department === "bihon"
          ? supabase
              .from("bh_cooking")
              .select("item_code, weight")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "bihon"
          ? supabase
              .from("bh_packing")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "bihon"
          ? supabase
              .from("bh_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "bihon"
          ? supabase
              .from("bihon_sku")
              .select("item_code, item_description, uom")
          : Promise.resolve({ data: [] }),

        // --- Snackfood blocks ---
        isAll || department === "snackfood"
          ? supabase
              .from("sf_blending")
              .select("item_code, usage")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase
              .from("sf_premix")
              .select("item_code, usage")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase
              .from("sf_mix")
              .select("item_code, weight")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase
              .from("sf_frying")
              .select("item_code, weight")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase
              .from("sf_flavoring")
              .select("item_code, weight")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase
              .from("sf_piece")
              .select("item_code, pcs")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase
              .from("sf_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "snackfood"
          ? supabase.from("sf_sku").select("item_code, item_description, uom")
          : Promise.resolve({ data: [] }),

        // --- Catmon blocks ---
        isAll || department === "catmon"
          ? supabase
              .from("catmon_mixing")
              .select("item_code, weight")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "catmon"
          ? supabase
              .from("catmon_frying_drying")
              .select("item_code, weight")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "catmon"
          ? supabase
              .from("catmon_packing")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "catmon"
          ? supabase
              .from("catmon_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "catmon"
          ? supabase
              .from("catmon_sku")
              .select("item_code, item_description, uom")
          : Promise.resolve({ data: [] }),

        // --- Kingsforth Sotanghon / Shared Core blocks ---
        isAll || department === "kf_sotanghon"
          ? supabase
              .from("kf_packing")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "kf_sotanghon"
          ? supabase
              .from("kf_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "kf_sotanghon"
          ? supabase
              .from("kf_seasoning")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department.startsWith("kf_")
          ? supabase.from("kf_sku").select("item_code, item_description, uom")
          : Promise.resolve({ data: [] }),

        // --- Kingsforth Hobe Express blocks ---
        isAll || department === "kf_hobe"
          ? supabase
              .from("kf_he_packing")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "kf_hobe"
          ? supabase
              .from("kf_he_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),

        // --- Kingsforth Canton blocks ---
        isAll || department === "kf_canton"
          ? supabase
              .from("kf_canton_packing")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "kf_canton"
          ? supabase
              .from("kf_canton_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),

        // --- Kingsforth Snackfood blocks ---
        isAll || department === "kf_sf"
          ? supabase
              .from("kf_sf_packing")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
        isAll || department === "kf_sf"
          ? supabase
              .from("kf_sf_fg")
              .select("item_code, qty")
              .in("prod_id", generatedProdIds)
          : Promise.resolve({ data: [] }),
      ])

      // --- Map Specifications and UOM Profiles ---
      bhSku.data?.forEach((s) =>
        globalDescMap.set(s.item_code, {
          desc: s.item_description,
          uom: s.uom || "pcs",
        })
      )
      sfSku.data?.forEach((s) =>
        globalDescMap.set(s.item_code, {
          desc: s.item_description,
          uom: s.uom || "kg",
        })
      )
      cmSku.data?.forEach((s) =>
        globalDescMap.set(s.item_code, {
          desc: s.item_description,
          uom: s.uom || "cs",
        })
      )
      kfSku.data?.forEach((s) =>
        globalDescMap.set(s.item_code, {
          desc: s.item_description,
          uom: s.uom || "cs",
        })
      )

      // Ensure SKU structures exist inside memory maps
      bhSku.data?.forEach((s) => getOrCreateRow(s.item_code))
      sfSku.data?.forEach((s) => getOrCreateRow(s.item_code))
      cmSku.data?.forEach((s) => getOrCreateRow(s.item_code))
      kfSku.data?.forEach((s) => getOrCreateRow(s.item_code))

      // --- Line Aggregation Implementations ---
      // Bihon
      bhCooking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.weight || 0)
      })
      bhPacking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      bhFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Snackfood
      sfBlend.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.usage || 0)
      })
      sfPremix.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.usage || 0)
      })
      sfMix.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.weight || 0)
      })
      sfFrying.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.weight || 0)
      })
      sfFlavor.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.weight || 0)
      })
      sfPiece.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.pcs || 0)
      })
      sfFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Catmon
      cmMix.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.weight || 0)
      })
      cmDry.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.weight || 0)
      })
      cmPacking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      cmFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Kingsforth Sotanghon
      kfPacking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      kfFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      kfSeasoning.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Kingsforth Hobe Express
      kfHePacking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      kfHeFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Kingsforth Canton
      kfCantonPacking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      kfCantonFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Kingsforth Snackfood
      kfSfPacking.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })
      kfSfFg.data?.forEach((r) => {
        getOrCreateRow(r.item_code).combined_value += Number(r.qty || 0)
      })

      // Filter down view to exclude zero activity lines
      const resultingRows = Array.from(localAggregationMap.values()).filter(
        (row) => row.combined_value > 0
      )

      setMergedRows(resultingRows)
    } catch (error: any) {
      console.error("Aggregation Pipeline Error:", error.message)
      alert(`Pipeline crash: ${error.message}`)
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="container mx-auto max-w-4xl space-y-6 p-6">
      <div className="flex items-center justify-between gap-4">
        <div className="flex flex-col gap-1">
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Production Audit Report
          </h1>
          <p className="text-xs text-slate-500">
            Horizontally collapsed yield profiles with database integrated
            units.
          </p>
        </div>

        {mergedRows && mergedRows.length > 0 && (
          <Button
            onClick={handleExportToExcel}
            variant="outline"
            className="h-9 border-slate-200 text-xs font-medium text-slate-700 shadow-xs hover:bg-slate-50 hover:text-slate-900"
          >
            <Download className="mr-1.5 h-3.5 w-3.5 text-slate-500" /> Export
            XLSX
          </Button>
        )}
      </div>

      <Card className="border-slate-200 bg-white shadow-xs">
        <CardContent className="p-4">
          <form
            onSubmit={handleGenerateReport}
            className="grid grid-cols-1 items-start gap-4 sm:grid-cols-4"
          >
            <div className="space-y-1.5">
              <Label
                htmlFor="department"
                className="text-xs font-bold text-slate-500 uppercase"
              >
                Department Line
              </Label>
              <Select value={department} onValueChange={setDepartment}>
                <SelectTrigger id="department" className="h-9 bg-white">
                  <SelectValue placeholder="Select Line Division" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All Departments</SelectItem>
                  <SelectItem value="bihon">Bihon</SelectItem>
                  <SelectItem value="snackfood">Snackfood</SelectItem>
                  <SelectItem value="catmon">Catmon</SelectItem>
                  <SelectItem value="kf_sotanghon">KF Sotanghon</SelectItem>
                  <SelectItem value="kf_hobe">KF Hobe Express</SelectItem>
                  <SelectItem value="kf_canton">KF Canton</SelectItem>
                  <SelectItem value="kf_sf">KF Snackfood</SelectItem>
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-1.5">
              <Label className="text-xs font-bold text-slate-500 uppercase">
                Operational Target Scope
              </Label>
              <div className="flex flex-col items-stretch gap-1.5">
                <div>
                  <span className="block text-xs text-slate-400">From</span>
                  <Popover>
                    <PopoverTrigger asChild>
                      <Button
                        variant="outline"
                        className="h-9 w-full justify-start text-xs font-normal"
                      >
                        <CalendarIcon className="mr-1.5 h-3.5 w-3.5 text-slate-400" />
                        <span className="truncate">
                          {startDate ? format(startDate, "MM/dd/yy") : "Start"}
                        </span>
                      </Button>
                    </PopoverTrigger>
                    <PopoverContent className="w-auto p-0" align="start">
                      <Calendar
                        mode="single"
                        selected={startDate}
                        onSelect={setStartDate}
                      />
                    </PopoverContent>
                  </Popover>
                </div>
                <div>
                  <span className="block text-xs text-slate-400">To</span>
                  <Popover>
                    <PopoverTrigger asChild>
                      <Button
                        variant="outline"
                        className="h-9 w-full justify-start text-xs font-normal"
                      >
                        <CalendarIcon className="mr-1.5 h-3.5 w-3.5 text-slate-400" />
                        <span className="truncate">
                          {endDate ? format(endDate, "MM/dd/yy") : "End"}
                        </span>
                      </Button>
                    </PopoverTrigger>
                    <PopoverContent className="w-auto p-0" align="start">
                      <Calendar
                        mode="single"
                        selected={endDate}
                        onSelect={setEndDate}
                        disabled={(date) => !!startDate && date < startDate}
                      />
                    </PopoverContent>
                  </Popover>
                </div>
              </div>
            </div>

            <div className="space-y-1.5">
              <Label className="text-xs font-bold text-slate-500 uppercase">
                Active Shifts
              </Label>
              <div className="flex h-9 items-center gap-4">
                <label className="flex cursor-pointer items-center space-x-2 text-xs font-medium">
                  <Checkbox
                    checked={shifts.day}
                    onCheckedChange={(c) =>
                      setShifts((p) => ({ ...p, day: !!c }))
                    }
                  />
                  <span>Day</span>
                </label>
                <label className="flex cursor-pointer items-center space-x-2 text-xs font-medium">
                  <Checkbox
                    checked={shifts.night}
                    onCheckedChange={(c) =>
                      setShifts((p) => ({ ...p, night: !!c }))
                    }
                  />
                  <span>Night</span>
                </label>
              </div>
            </div>

            <Button
              type="submit"
              className="h-9 bg-blue-600 text-xs font-medium text-white hover:bg-blue-700"
              disabled={
                isLoading ||
                !department ||
                (!shifts.day && !shifts.night) ||
                !startDate ||
                !endDate
              }
            >
              {isLoading ? (
                <>
                  <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" />{" "}
                  Summing...
                </>
              ) : (
                <>
                  <Search className="mr-1.5 h-3.5 w-3.5" /> Compile Line
                </>
              )}
            </Button>
          </form>
        </CardContent>
      </Card>

      {/* RENDER FLAT INTERFACE VIEW */}
      {!mergedRows ? (
        <div className="rounded-xl border border-dashed bg-slate-50/50 p-12 text-center text-slate-400">
          <FileSpreadsheet className="mx-auto mb-2 h-8 w-8 text-slate-300" />
          <p className="text-xs font-medium">
            Select a department line to extract combined metrics.
          </p>
        </div>
      ) : mergedRows.length === 0 ? (
        <div className="rounded-xl border bg-white p-12 text-center text-slate-500">
          <p className="text-xs font-semibold">
            No data points logged inside requested search criteria.
          </p>
        </div>
      ) : (
        <Card className="overflow-hidden border-slate-200 bg-white shadow-sm">
          <CardHeader className="border-b border-slate-100 bg-slate-50 px-4 py-3">
            <CardTitle className="text-xs font-bold tracking-wider text-slate-800 uppercase">
              Yield Profile:{" "}
              {department === "all"
                ? "All Divisions Consolidation"
                : department.replace("kf_", "Kingsforth ")}
            </CardTitle>
          </CardHeader>
          <CardContent className="p-0">
            <Table>
              <TableHeader className="bg-slate-100/70">
                <TableRow className="hover:bg-transparent">
                  <TableHead className="w-[180px] text-xs font-bold text-slate-700">
                    SKU Code
                  </TableHead>
                  <TableHead className="text-xs font-bold text-slate-700">
                    Item Description
                  </TableHead>
                  <TableHead className="w-[180px] text-right text-xs font-bold text-slate-700">
                    Total Extracted Value
                  </TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {mergedRows.map((row) => (
                  <TableRow
                    key={row.item_code}
                    className="transition-colors hover:bg-slate-50/60"
                  >
                    <TableCell className="font-mono text-xs font-bold text-slate-900">
                      {row.item_code}
                    </TableCell>
                    <TableCell
                      className="max-w-md truncate text-xs text-slate-600"
                      title={row.item_description}
                    >
                      {row.item_description}
                    </TableCell>
                    <TableCell className="text-right font-mono text-xs font-bold whitespace-nowrap text-blue-600">
                      {row.combined_value.toLocaleString()}{" "}
                      <span className="ml-1 text-[10px] font-medium tracking-wide text-slate-400 uppercase">
                        {row.uom}
                      </span>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}
    </div>
  )
}
