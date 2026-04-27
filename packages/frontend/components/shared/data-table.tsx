"use client";

import {
  type ColumnDef,
  type SortingState,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  getSortedRowModel,
  useReactTable,
} from "@tanstack/react-table";
import { ChevronDown, ChevronUp, ChevronsUpDown } from "lucide-react";
import { useState } from "react";

import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

interface DataTableProps<TData, TValue> {
  columns: ColumnDef<TData, TValue>[];
  data: TData[];
  loading?: boolean;
  pageSize?: number;
  globalFilter?: boolean;
  exportFileName?: string;
}

function SkeletonRow({ columns }: { columns: number }) {
  return (
    <tr>
      {Array.from({ length: columns }).map((_, i) => (
        <td key={i} className="px-4 py-3">
          <div className="h-4 animate-pulse rounded bg-cream-200" />
        </td>
      ))}
    </tr>
  );
}

export function DataTable<TData, TValue>({
  columns,
  data,
  loading = false,
  pageSize = 20,
  globalFilter = false,
}: DataTableProps<TData, TValue>) {
  const [sorting, setSorting] = useState<SortingState>([]);
  const [globalFilterValue, setGlobalFilterValue] = useState("");

  const table = useReactTable({
    data,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
    onSortingChange: setSorting,
    state: {
      sorting,
      globalFilter: globalFilterValue,
    },
    onGlobalFilterChange: setGlobalFilterValue,
    initialState: {
      pagination: { pageSize },
    },
  });

  return (
    <div className="flex flex-col gap-3">
      {globalFilter && (
        <input
          type="search"
          value={globalFilterValue}
          onChange={(e) => setGlobalFilterValue(e.target.value)}
          placeholder="検索..."
          className="w-full max-w-xs rounded-xl border border-cream-300 bg-cream-50 px-3 py-1.5 text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
        />
      )}

      <div className="overflow-x-auto rounded-xl border border-cream-300">
        <table className="w-full text-sm">
          <thead className="bg-cream-100">
            {table.getHeaderGroups().map((headerGroup) => (
              <tr key={headerGroup.id} className="border-b border-cream-300">
                {headerGroup.headers.map((header) => {
                  const canSort = header.column.getCanSort();
                  const sortDir = header.column.getIsSorted();

                  return (
                    <th
                      key={header.id}
                      className={cn(
                        "px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-ink-muted",
                        canSort && "cursor-pointer select-none hover:text-ink",
                      )}
                      onClick={canSort ? header.column.getToggleSortingHandler() : undefined}
                    >
                      <span className="flex items-center gap-1">
                        {flexRender(header.column.columnDef.header, header.getContext())}
                        {canSort && (
                          <span className="text-ink-muted">
                            {sortDir === "asc" ? (
                              <ChevronUp className="h-3.5 w-3.5" />
                            ) : sortDir === "desc" ? (
                              <ChevronDown className="h-3.5 w-3.5" />
                            ) : (
                              <ChevronsUpDown className="h-3.5 w-3.5 opacity-40" />
                            )}
                          </span>
                        )}
                      </span>
                    </th>
                  );
                })}
              </tr>
            ))}
          </thead>
          <tbody className="divide-y divide-cream-200 bg-white">
            {loading ? (
              Array.from({ length: pageSize > 5 ? 5 : pageSize }).map((_, i) => (
                <SkeletonRow key={i} columns={columns.length} />
              ))
            ) : table.getRowModel().rows.length === 0 ? (
              <tr>
                <td colSpan={columns.length} className="px-4 py-8 text-center text-ink-muted">
                  データがありません
                </td>
              </tr>
            ) : (
              table.getRowModel().rows.map((row) => (
                <tr key={row.id} className="transition-colors hover:bg-cream-50">
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id} className="px-4 py-2.5 text-ink">
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      <div className="flex items-center justify-between gap-4 text-sm text-ink-muted">
        <span>
          {table.getFilteredRowModel().rows.length} 件中{" "}
          {table.getState().pagination.pageIndex * table.getState().pagination.pageSize + 1}〜
          {Math.min(
            (table.getState().pagination.pageIndex + 1) * table.getState().pagination.pageSize,
            table.getFilteredRowModel().rows.length,
          )}{" "}
          件表示
        </span>
        <div className="flex items-center gap-2">
          <select
            value={table.getState().pagination.pageSize}
            onChange={(e) => table.setPageSize(Number(e.target.value))}
            className="rounded-lg border border-cream-300 bg-cream-50 px-2 py-1 text-sm text-ink focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          >
            {[10, 20, 50, 100].map((size) => (
              <option key={size} value={size}>
                {size}件
              </option>
            ))}
          </select>
          <Button
            variant="outline"
            size="sm"
            onClick={() => table.previousPage()}
            disabled={!table.getCanPreviousPage()}
          >
            前へ
          </Button>
          <span className="text-xs">
            {table.getState().pagination.pageIndex + 1} / {table.getPageCount()}
          </span>
          <Button
            variant="outline"
            size="sm"
            onClick={() => table.nextPage()}
            disabled={!table.getCanNextPage()}
          >
            次へ
          </Button>
        </div>
      </div>
    </div>
  );
}
