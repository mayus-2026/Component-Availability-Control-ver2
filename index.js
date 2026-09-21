(function () {
  "use strict";

  var ROWS = [];
  var LEVEL1_ROWS = [];
  var AS_OF = stripTime(new Date());
  var CUR_Y = AS_OF.getFullYear(),
    CUR_M = AS_OF.getMonth();
  var downloadsCap = null;
  var FG_AGG = [];
  var FG_AGG_BY_CODE = {};
  var STOCK_INDEX = {}; // key: component|plant|level -> {stock, grossDemand, netShortfall, lines, fgSet}
  var SHORT_ROWS_BY_COMP = {};
  var CURRENT_PLANT = ""; // '' = all plants
  var ALL_PLANTS = [];

  var EXPECTED = [
    "plant",
    "company_code",
    "fg_material",
    "fg_description",
    "bom_level",
    "parent_material",
    "bom_item_number",
    "bom_item_path",
    "component_material",
    "component_description",
    "sales_order",
    "item_number",
    "schedule_line",
    "requirement_date",
    "fg_demand",
    "bom_qty_per_fg",
    "gross_child_part_demand_to_cover_fg",
    "gross_child_part_stock",
    "gross_shortage_or_excess",
    "gross_status_flag",
    "net_demand_this_line",
    "net_coverage_flag",
    "bom_number",
    "component_uom",
  ];
  var REQUIRED = [
    "plant",
    "fg_material",
    "component_material",
    "bom_level",
    "requirement_date",
    "net_coverage_flag",
    "net_demand_this_line",
  ];

  function stripTime(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  function norm(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .replace(/[\s]+/g, "_");
  }
  function fmtInt(n) {
    return Math.round(n).toLocaleString("en-US");
  }
  function fmtNum(n) {
    return (Math.round(n * 100) / 100).toLocaleString("en-US", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2,
    });
  }
  function fmtPct(n) {
    return Math.round(n * 10) / 10 + "%";
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function fmtDate(d) {
    return d
      ? d.toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "short",
          year: "numeric",
        })
      : "&mdash;";
  }
  function debounce(fn, ms) {
    var t;
    return function () {
      clearTimeout(t);
      var a = arguments;
      t = setTimeout(function () {
        fn.apply(null, a);
      }, ms);
    };
  }

  function monthDiff(reqDate) {
    // positive = requirement month is N months in the past relative to current month
    if (!reqDate) return null;
    return (
      CUR_Y * 12 + CUR_M - (reqDate.getFullYear() * 12 + reqDate.getMonth())
    );
  }

  function coverageBucket(flag) {
    var f = String(flag || "").toUpperCase();
    if (f.indexOf("NOT COVERED") !== -1) return "NOT COVERED";
    if (f.indexOf("PARTIAL") !== -1) return "PARTIALLY";
    if (f.indexOf("FULLY") !== -1 || f.indexOf("COVERED") !== -1)
      return "FULLY";
    return "UNKNOWN";
  }
  function pillClass(bucket) {
    if (bucket === "NOT COVERED") return "red";
    if (bucket === "PARTIALLY") return "amber";
    if (bucket === "FULLY") return "green";
    return "green";
  }
  function pillHtml(bucket) {
    return (
      '<span class="pill ' + pillClass(bucket) + '">' + esc(bucket) + "</span>"
    );
  }
  function isShortBucket(b) {
    return b === "NOT COVERED" || b === "PARTIALLY";
  }

  function rowsForPlant(rows) {
    return CURRENT_PLANT
      ? rows.filter(function (r) {
          return r.plant === CURRENT_PLANT;
        })
      : rows;
  }

  /* ---------------- file loading ---------------- */
  var dropzone = document.getElementById("dropzone");
  var fileInput = document.getElementById("fileInput");
  var parseStatus = document.getElementById("parseStatus");

  dropzone.addEventListener("click", function () {
    fileInput.click();
  });
  dropzone.addEventListener("dragover", function (e) {
    e.preventDefault();
    dropzone.classList.add("drag");
  });
  dropzone.addEventListener("dragleave", function () {
    dropzone.classList.remove("drag");
  });
  dropzone.addEventListener("drop", function (e) {
    e.preventDefault();
    dropzone.classList.remove("drag");
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener("change", function () {
    if (fileInput.files.length) handleFile(fileInput.files[0]);
  });
  document.getElementById("loadNewBtn").addEventListener("click", function () {
    document.getElementById("app").style.display = "none";
    document.getElementById("landing").style.display = "block";
    document.getElementById("fileTag").style.display = "none";
    document.getElementById("loadNewBtn").style.display = "none";
    fileInput.value = "";
    parseStatus.textContent = "";
    parseStatus.className = "";
  });

  function handleFile(file) {
    parseStatus.className = "";
    parseStatus.textContent = "Reading " + file.name + "\u2026";
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = XLSX.read(data, { type: "array", cellDates: true });
        var sheetName =
          wb.SheetNames.indexOf("Query result") !== -1
            ? "Query result"
            : wb.SheetNames[0];
        var ws = wb.Sheets[sheetName];
        var rows2d = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          raw: true,
          defval: null,
        });
        if (!rows2d.length) {
          throw new Error("The sheet appears to be empty.");
        }
        var header = rows2d[0].map(norm);
        var COLMAP = {};
        header.forEach(function (h, idx) {
          COLMAP[h] = idx;
        });
        var missing = REQUIRED.filter(function (c) {
          return !(c in COLMAP);
        });
        if (missing.length) {
          throw new Error(
            "Missing required column(s): " +
              missing.join(", ") +
              ". Found columns: " +
              header.join(", "),
          );
        }
        ROWS = [];
        for (var i = 1; i < rows2d.length; i++) {
          var r = rows2d[i];
          if (
            !r ||
            r.every(function (v) {
              return v === null || v === "";
            })
          )
            continue;
          var obj = {};
          EXPECTED.forEach(function (c) {
            var idx = COLMAP[c];
            obj[c] = idx === undefined ? null : r[idx];
          });
          obj.bom_level = Number(obj.bom_level) || 0;
          obj.fg_demand = Number(obj.fg_demand) || 0;
          obj.bom_qty_per_fg = Number(obj.bom_qty_per_fg) || 0;
          obj.gross_child_part_demand_to_cover_fg =
            Number(obj.gross_child_part_demand_to_cover_fg) || 0;
          obj.gross_child_part_stock = Number(obj.gross_child_part_stock) || 0;
          obj.gross_shortage_or_excess =
            Number(obj.gross_shortage_or_excess) || 0;
          obj.net_demand_this_line = Number(obj.net_demand_this_line) || 0;
          obj.sales_order =
            obj.sales_order == null ? "" : String(obj.sales_order).trim();
          obj.item_number =
            obj.item_number == null ? "" : String(obj.item_number).trim();
          obj.schedule_line =
            obj.schedule_line == null ? "" : String(obj.schedule_line).trim();
          obj.plant = obj.plant == null ? "" : String(obj.plant).trim();
          obj.company_code =
            obj.company_code == null ? "" : String(obj.company_code).trim();
          obj.fg_material =
            obj.fg_material == null ? "" : String(obj.fg_material).trim();
          obj.fg_description =
            obj.fg_description == null ? "" : String(obj.fg_description).trim();
          obj.component_material =
            obj.component_material == null
              ? ""
              : String(obj.component_material).trim();
          obj.component_description =
            obj.component_description == null
              ? ""
              : String(obj.component_description).trim();
          obj.bom_item_path =
            obj.bom_item_path == null ? "" : String(obj.bom_item_path).trim();
          var rd = obj.requirement_date;
          if (rd instanceof Date && !isNaN(rd)) {
            obj.requirement_date = stripTime(rd);
          } else if (typeof rd === "number") {
            obj.requirement_date = stripTime(
              new Date(Math.round((rd - 25569) * 86400 * 1000)),
            );
          } else if (typeof rd === "string" && rd) {
            var pd = new Date(rd);
            obj.requirement_date = isNaN(pd) ? null : stripTime(pd);
          } else {
            obj.requirement_date = null;
          }
          obj._covBucket = coverageBucket(obj.net_coverage_flag);
          obj._monthDiff = monthDiff(obj.requirement_date);
          obj._key = [
            obj.plant,
            obj.component_material,
            obj.sales_order,
            obj.item_number,
            obj.schedule_line,
            obj.bom_item_path,
          ].join("||");
          obj._sortKey = [
            obj.requirement_date ? obj.requirement_date.getTime() : 9e15,
            obj.sales_order,
            obj.item_number,
            obj.schedule_line,
            obj.bom_item_path,
          ].join("~");
          ROWS.push(obj);
        }
        if (!ROWS.length)
          throw new Error("No data rows found after the header.");
        LEVEL1_ROWS = ROWS.filter(function (r) {
          return r.bom_level === 1;
        });
        var plantSet = {};
        ROWS.forEach(function (r) {
          if (r.plant) plantSet[r.plant] = 1;
        });
        ALL_PLANTS = Object.keys(plantSet).sort();
        CURRENT_PLANT = "";

        parseStatus.textContent = "Parsed " + fmtInt(ROWS.length) + " rows.";
        document.getElementById("fileTag").textContent =
          file.name + "  \u2022  " + fmtInt(ROWS.length) + " rows";
        document.getElementById("fileTag").style.display = "inline-block";
        document.getElementById("loadNewBtn").style.display = "inline-block";

        setupOneTimeUI();
        refreshAll();

        document.getElementById("landing").style.display = "none";
        document.getElementById("app").style.display = "block";
      } catch (err) {
        parseStatus.className = "err";
        parseStatus.textContent = "Could not read this file: " + err.message;
      }
    };
    reader.onerror = function () {
      parseStatus.className = "err";
      parseStatus.textContent = "Could not read the file from disk.";
    };
    reader.readAsArrayBuffer(file);
  }

  /* ---------------- tabs ---------------- */
  document.querySelectorAll("nav.tabs button").forEach(function (btn) {
    btn.addEventListener("click", function () {
      activateTab(btn.dataset.tab);
    });
  });
  function activateTab(tab) {
    document.querySelectorAll("nav.tabs button").forEach(function (b) {
      b.classList.toggle("active", b.dataset.tab === tab);
    });
    document.querySelectorAll(".panel").forEach(function (p) {
      p.classList.remove("active");
    });
    document.getElementById("panel-" + tab).classList.add("active");
  }

  var oneTimeUIDone = false;
  function setupOneTimeUI() {
    var gp = document.getElementById("globalPlantFilter");
    gp.innerHTML =
      '<option value="">All plants (combined)</option>' +
      ALL_PLANTS.map(function (p) {
        return '<option value="' + esc(p) + '">' + esc(p) + "</option>";
      }).join("");
    if (oneTimeUIDone) return;
    oneTimeUIDone = true;
    gp.addEventListener("change", function () {
      CURRENT_PLANT = gp.value;
      refreshAll();
    });
    document
      .getElementById("plantBarClear")
      .addEventListener("click", function () {
        CURRENT_PLANT = "";
        gp.value = "";
        refreshAll();
      });
    document.getElementById("fgSearch").addEventListener(
      "input",
      debounce(function () {
        fgFocusState.page = 0;
        renderFgList();
      }, 200),
    );
    document.getElementById("fgSort").addEventListener("change", function () {
      fgFocusState.page = 0;
      renderFgList();
    });
    document
      .getElementById("fgPrevPage")
      .addEventListener("click", function () {
        if (fgFocusState.page > 0) {
          fgFocusState.page--;
          renderFgList();
        }
      });
    document
      .getElementById("fgNextPage")
      .addEventListener("click", function () {
        fgFocusState.page++;
        renderFgList();
      });
    document
      .getElementById("fgDetailClose")
      .addEventListener("click", function () {
        document.getElementById("fgDetailCard").style.display = "none";
        fgFocusState.selected = null;
      });
    document.getElementById("ppScope").addEventListener("change", function () {
      ppState.page = 0;
      renderPpTable();
    });
    document
      .getElementById("ppCoverage")
      .addEventListener("change", function () {
        ppState.page = 0;
        renderPpTable();
      });
    document.getElementById("ppSearch").addEventListener(
      "input",
      debounce(function () {
        ppState.page = 0;
        renderPpTable();
      }, 200),
    );
    document
      .getElementById("ppPrevPage")
      .addEventListener("click", function () {
        if (ppState.page > 0) {
          ppState.page--;
          renderPpTable();
        }
      });
    document
      .getElementById("ppNextPage")
      .addEventListener("click", function () {
        ppState.page++;
        renderPpTable();
      });
    document.getElementById("svLevel").addEventListener("change", function () {
      svState.page = 0;
      renderStockVis();
    });
    document.getElementById("svShow").addEventListener("change", function () {
      svState.page = 0;
      renderStockVis();
    });
    document.getElementById("svSearch").addEventListener(
      "input",
      debounce(function () {
        svState.page = 0;
        renderStockVis();
      }, 200),
    );
    document
      .getElementById("svPrevPage")
      .addEventListener("click", function () {
        if (svState.page > 0) {
          svState.page--;
          renderStockVis();
        }
      });
    document
      .getElementById("svNextPage")
      .addEventListener("click", function () {
        svState.page++;
        renderStockVis();
      });
    document
      .getElementById("simCompSearch")
      .addEventListener("input", debounce(refreshSimCompOptions, 200));
    document
      .getElementById("simCompSelect")
      .addEventListener("change", refreshSimPlantLevelOptions);
    document
      .getElementById("simRunBtn")
      .addEventListener("click", runSimulation);
    ["fPlant", "fLevel", "fCoverage"].forEach(function (id) {
      document.getElementById(id).addEventListener("change", function () {
        explorerState.page = 0;
        renderExplorer();
      });
    });
    document.getElementById("fMinQty").addEventListener(
      "input",
      debounce(function () {
        explorerState.page = 0;
        renderExplorer();
      }, 200),
    );
    document.getElementById("fSearch").addEventListener(
      "input",
      debounce(function () {
        explorerState.page = 0;
        renderExplorer();
      }, 200),
    );
    document.getElementById("prevPage").addEventListener("click", function () {
      if (explorerState.page > 0) {
        explorerState.page--;
        renderExplorer();
      }
    });
    document.getElementById("nextPage").addEventListener("click", function () {
      explorerState.page++;
      renderExplorer();
    });
    document.getElementById("exportBtn").addEventListener("click", exportCSV);
  }

  /* ================= per-FG aggregation (level 1 only) ================= */
  function buildFgAgg(rows) {
    var map = {};
    var seenQtyKeys = {};
    rows.forEach(function (r) {
      var k = r.fg_material;
      if (!k) return;
      if (!map[k]) {
        map[k] = {
          fg_material: k,
          fg_description: r.fg_description,
          plantSet: {},
          soSet: {},
          totalOpenQty: 0,
          componentsTotal: 0,
          componentsShort: 0,
          componentsPartial: 0,
          earliest: null,
          worstComponent: null,
          worstShortfall: -1,
        };
      }
      var o = map[k];
      if (r.plant) o.plantSet[r.plant] = 1;
      if (r.sales_order) o.soSet[r.sales_order] = 1;
      var qk =
        k + "|" + r.sales_order + "|" + r.item_number + "|" + r.schedule_line;
      if (!seenQtyKeys[qk]) {
        seenQtyKeys[qk] = 1;
        o.totalOpenQty += r.fg_demand;
      }
      o.componentsTotal++;
      if (r._covBucket === "NOT COVERED") o.componentsShort++;
      if (r._covBucket === "PARTIALLY") o.componentsPartial++;
      if (
        r.requirement_date &&
        (!o.earliest || r.requirement_date < o.earliest)
      )
        o.earliest = r.requirement_date;
      if (r.net_demand_this_line > o.worstShortfall) {
        o.worstShortfall = r.net_demand_this_line;
        o.worstComponent = r.component_material;
      }
    });
    var arr = Object.keys(map).map(function (k) {
      var o = map[k];
      o.plants = Object.keys(o.plantSet);
      o.soCount = Object.keys(o.soSet).length;
      o.readiness = o.componentsTotal
        ? Math.round(
            (1 -
              (o.componentsShort + o.componentsPartial) / o.componentsTotal) *
              100,
          )
        : 100;
      return o;
    });
    return arr;
  }

  /* ================= stock index: key = component|plant|level ================= */
  function buildStockIndex(rows) {
    var groups = {};
    rows.forEach(function (r) {
      var key = r.component_material + "|" + r.plant + "|" + r.bom_level;
      if (!groups[key]) groups[key] = [];
      groups[key].push(r);
    });
    var idx = {};
    Object.keys(groups).forEach(function (key) {
      var g = groups[key];
      g.sort(function (a, b) {
        return a._sortKey < b._sortKey ? -1 : a._sortKey > b._sortKey ? 1 : 0;
      });
      var first = g[0];
      var grossDemand = 0,
        netShortfall = 0,
        fgSet = {};
      g.forEach(function (r) {
        grossDemand += r.gross_child_part_demand_to_cover_fg;
        netShortfall += r.net_demand_this_line;
        if (r.fg_material) fgSet[r.fg_material] = 1;
      });
      idx[key] = {
        component_material: first.component_material,
        component_description: first.component_description,
        uom: first.component_uom,
        plant: first.plant,
        level: first.bom_level,
        stock: first.gross_child_part_stock,
        grossDemand: grossDemand,
        netShortfall: netShortfall,
        lines: g.length,
        fgCount: Object.keys(fgSet).length,
      };
    });
    return idx;
  }
  function getStock(component, plant, level) {
    var v = STOCK_INDEX[component + "|" + plant + "|" + level];
    return v ? v.stock : null;
  }

  /* ================= MAIN REFRESH (re-run whenever plant filter changes) ================= */
  function refreshAll() {
    var ROWS_S = rowsForPlant(ROWS);
    var L1_S = rowsForPlant(LEVEL1_ROWS);
    FG_AGG = buildFgAgg(L1_S);
    FG_AGG_BY_CODE = {};
    FG_AGG.forEach(function (o) {
      FG_AGG_BY_CODE[o.fg_material] = o;
    });
    STOCK_INDEX = buildStockIndex(ROWS_S);
    buildShortIndex(ROWS_S);

    var pb = document.getElementById("plantBar");
    var note = document.getElementById("plantBarNote");
    var clearBtn = document.getElementById("plantBarClear");
    if (CURRENT_PLANT) {
      pb.classList.add("active");
      note.textContent =
        "Every tab below (except Plant Breakdown) is scoped to plant " +
        CURRENT_PLANT +
        " only.";
      clearBtn.style.display = "inline-block";
    } else {
      pb.classList.remove("active");
      note.textContent =
        ALL_PLANTS.length +
        " plants combined. Pick a plant above to focus every tab on just that plant.";
      clearBtn.style.display = "none";
    }

    renderOverview(ROWS_S);
    renderFgList();
    if (fgFocusState.selected && FG_AGG_BY_CODE[fgFocusState.selected])
      focusFg(fgFocusState.selected);
    else {
      document.getElementById("fgDetailCard").style.display = "none";
    }
    renderPpTable();
    renderStockVis();
    refreshSimCompOptions();
    renderTimeChart(ROWS_S);
    renderPastDueTable(ROWS_S);
    renderPlantTable();
    renderIntegrity();
    renderScopeDetail();
    setupExplorerOptions();
    renderExplorer();
  }

  function renderOverview(ROWS_S) {
    document.getElementById("asOfLabel").textContent = fmtDate(AS_OF);
    var totalRows = ROWS_S.length;
    var soSet = {},
      fgSet = {},
      compSet = {},
      plantSet = {};
    var covCount = { FULLY: 0, PARTIALLY: 0, "NOT COVERED": 0, UNKNOWN: 0 };
    var levelCount = {};
    var oldDemandCount = 0,
      pastDueCount = 0,
      shortRows = 0,
      pastDueSOset = {},
      oldSOset = {};

    ROWS_S.forEach(function (r) {
      if (r.sales_order) soSet[r.sales_order] = 1;
      if (r.fg_material) fgSet[r.fg_material] = 1;
      if (r.component_material) compSet[r.component_material] = 1;
      if (r.plant) plantSet[r.plant] = 1;
      covCount[r._covBucket] = (covCount[r._covBucket] || 0) + 1;
      levelCount[r.bom_level] = (levelCount[r.bom_level] || 0) + 1;
      if (isShortBucket(r._covBucket)) {
        shortRows++;
        if (r._monthDiff !== null && r._monthDiff >= 1) {
          pastDueCount++;
          if (r.sales_order) pastDueSOset[r.sales_order] = 1;
          if (r._monthDiff >= 3) {
            oldDemandCount++;
            if (r.sales_order) oldSOset[r.sales_order] = 1;
          }
        }
      }
    });
    var soCount = Object.keys(soSet).length,
      fgCount = Object.keys(fgSet).length,
      compCount = Object.keys(compSet).length,
      plantCount = Object.keys(plantSet).length;
    var notCoveredSO = {};
    ROWS_S.forEach(function (r) {
      if (r._covBucket === "NOT COVERED" && r.sales_order)
        notCoveredSO[r.sales_order] = 1;
    });
    var notCoveredSOCount = Object.keys(notCoveredSO).length;
    var fgBlockedCount = FG_AGG.filter(function (o) {
      return o.componentsShort > 0;
    }).length;

    var banners = [];
    if (!CURRENT_PLANT && ALL_PLANTS.length > 1) {
      banners.push({
        cls: "amber",
        icon: "\u26A0",
        html:
          "<b>Multi-plant extract</b> &mdash; combining <b>" +
          ALL_PLANTS.length +
          " plants</b>. Use the plant selector above to focus on one plant at a time; most procurement and completion decisions are plant-specific.",
      });
    }
    if (oldDemandCount > 0) {
      banners.push({
        cls: "purple",
        icon: "\u2753",
        html:
          "<b>" +
          fmtInt(oldDemandCount) +
          " lines are 3+ months past their requirement month</b>, across " +
          fmtInt(Object.keys(oldSOset).length) +
          " sales orders. These are unlikely to be live supply-chain risk &mdash; verify whether they are still valid open orders or should be cancelled/cleared in SAP.",
      });
    }
    var pastDuePct = shortRows ? (pastDueCount / shortRows) * 100 : 0;
    if (pastDuePct > 40 && shortRows > 0) {
      banners.push({
        cls: "red",
        icon: "\u2716",
        html:
          "<b>" +
          fmtPct(pastDuePct) +
          " of uncovered demand is past its requirement month</b> (" +
          fmtInt(pastDueCount) +
          " of " +
          fmtInt(shortRows) +
          " short lines) &mdash; this is aged open backlog, not current-month build risk.",
      });
    }
    document.getElementById("topBanners").innerHTML = banners
      .map(function (b) {
        return (
          '<div class="banner ' +
          b.cls +
          '"><span class="b-icon">' +
          b.icon +
          "</span><div>" +
          b.html +
          "</div></div>"
        );
      })
      .join("");

    var kpis = [
      {
        label: "Demand lines" + (CURRENT_PLANT ? " (this plant)" : ""),
        val: fmtInt(totalRows),
        sub:
          fmtInt(plantCount) +
          " plant(s) \u00b7 " +
          fmtInt(fgCount) +
          " FG materials",
        cls: "",
      },
      {
        label: "FGs blocked by a shortage",
        val: fmtInt(fgBlockedCount),
        sub:
          fmtPct(fgCount ? (fgBlockedCount / fgCount) * 100 : 0) +
          " of FGs in scope",
        cls: fgBlockedCount > 0 ? "bad" : "good",
      },
      {
        label: "Sales orders in scope",
        val: fmtInt(soCount),
        sub: fmtInt(notCoveredSOCount) + " with uncovered demand",
        cls: notCoveredSOCount > 0 ? "warn" : "good",
      },
      {
        label: "Components in scope",
        val: fmtInt(compCount),
        sub: "across " + fmtInt(Object.keys(levelCount).length) + " BOM levels",
        cls: "",
      },
      {
        label: "Net not covered",
        val: fmtPct(
          totalRows ? (covCount["NOT COVERED"] / totalRows) * 100 : 0,
        ),
        sub:
          fmtInt(covCount["NOT COVERED"]) +
          " of " +
          fmtInt(totalRows) +
          " lines",
        cls: covCount["NOT COVERED"] > 0 ? "bad" : "good",
      },
      {
        label: "Past-due (prior months)",
        val: fmtPct(pastDuePct),
        sub: fmtInt(pastDueCount) + " short lines, prior months",
        cls: pastDuePct > 25 ? "bad" : pastDuePct > 0 ? "warn" : "good",
      },
      {
        label: "Old demand (3+ months)",
        val: fmtInt(oldDemandCount),
        sub: "possible SAP data-cleanup item",
        cls: oldDemandCount > 0 ? "old" : "good",
      },
    ];
    document.getElementById("kpiRow").innerHTML = kpis
      .map(function (k) {
        return (
          '<div class="kpi ' +
          k.cls +
          '"><div class="k-label">' +
          k.label +
          '</div><div class="k-val">' +
          k.val +
          '</div><div class="k-sub">' +
          k.sub +
          "</div></div>"
        );
      })
      .join("");

    var topComp = computeComponentAgg("direct", "short", ROWS_S).sort(
      function (a, b) {
        return b.fgSet_n - a.fgSet_n;
      },
    )[0];
    var scopeLabel = CURRENT_PLANT
      ? "plant <b>" + esc(CURRENT_PLANT) + "</b>"
      : "<b>" + plantCount + "</b> plant(s) combined";
    var narrativeHtml =
      "<h3>Executive read" +
      (CURRENT_PLANT ? " \u2014 " + esc(CURRENT_PLANT) : "") +
      "</h3>";
    narrativeHtml +=
      "<p>This view covers <b>" +
      fmtInt(totalRows) +
      "</b> demand lines for " +
      scopeLabel +
      " and <b>" +
      fgCount +
      "</b> finished-good materials. <b>" +
      fmtInt(fgBlockedCount) +
      " FGs</b> currently cannot fully complete because at least one directly-consumed component is short.</p>";
    if (topComp) {
      var stockNote =
        topComp.plants.length === 1
          ? ", with " +
            fmtNum(
              getStock(topComp.component_material, topComp.plants[0], 1) || 0,
            ) +
            " units currently in stock against " +
            fmtNum(
              topComp.total +
                (getStock(topComp.component_material, topComp.plants[0], 1) ||
                  0),
            ) +
            " total gross demand"
          : "";
      narrativeHtml +=
        "<p>The single highest-leverage component right now is <b>" +
        esc(topComp.component_material) +
        "</b> (" +
        esc(topComp.component_description || "") +
        ") &mdash; it is blocking <b>" +
        topComp.fgSet_n +
        " FG materials</b> and <b>" +
        topComp.soSet_n +
        " sales orders</b> at once" +
        stockNote +
        ". See Procurement Priority and Supply Simulator for the full ranked list and a what-if calculator.</p>";
    }
    if (oldDemandCount > 0) {
      narrativeHtml +=
        "<p>Note: " +
        fmtInt(oldDemandCount) +
        ' short lines are 3+ months old. These are shown separately throughout this tool (Aging &amp; Delivery Risk tab) and are excluded from being called "current" risk &mdash; treat them as a data-hygiene item for planning/CS to confirm, not a build blocker for this month.</p>';
    }
    document.getElementById("narrative").innerHTML = narrativeHtml;

    renderDonutish(
      "coverageChart",
      [
        {
          label: "Fully covered",
          value: covCount["FULLY"] || 0,
          color: "#2E8B57",
        },
        {
          label: "Partially covered",
          value: covCount["PARTIALLY"] || 0,
          color: "#C67C1E",
        },
        {
          label: "Not covered",
          value: covCount["NOT COVERED"] || 0,
          color: "#B0362A",
        },
      ],
      totalRows,
    );
    var levelKeys = Object.keys(levelCount)
      .map(Number)
      .sort(function (a, b) {
        return a - b;
      });
    renderBarList(
      "levelChart",
      levelKeys.map(function (l) {
        return {
          label: "Level " + l,
          value: levelCount[l],
          color: l === 1 ? "#1C7293" : "#9AA7B4",
        };
      }),
    );
    renderSOTable(ROWS_S);
  }

  function renderBarList(elId, items) {
    var max =
      Math.max.apply(
        null,
        items.map(function (i) {
          return i.value;
        }),
      ) || 1;
    var html = '<div class="bar-chart">';
    items.forEach(function (i) {
      var pct = max ? (i.value / max) * 100 : 0;
      html +=
        '<div class="row"><div class="lbl">' +
        esc(i.label) +
        '</div><div class="track"><div class="fill" style="width:' +
        pct +
        "%;background:" +
        i.color +
        ';"></div></div><div class="val">' +
        fmtInt(i.value) +
        "</div></div>";
    });
    document.getElementById(elId).innerHTML = html + "</div>";
  }
  function renderDonutish(elId, items, total) {
    var html =
      '<div style="display:flex;height:26px;border-radius:4px;overflow:hidden;margin-bottom:14px;">';
    items.forEach(function (i) {
      var pct = total ? (i.value / total) * 100 : 0;
      if (pct > 0)
        html +=
          '<div style="width:' +
          pct +
          "%;background:" +
          i.color +
          ';" title="' +
          esc(i.label) +
          ": " +
          fmtInt(i.value) +
          '"></div>';
    });
    html += '</div><div class="bar-chart">';
    items.forEach(function (i) {
      var pct = total ? (i.value / total) * 100 : 0;
      html +=
        '<div class="row"><div class="lbl" style="display:flex;align-items:center;gap:6px;"><span style="width:9px;height:9px;border-radius:2px;background:' +
        i.color +
        ';display:inline-block;"></span>' +
        esc(i.label) +
        '</div><div class="track"><div class="fill" style="width:' +
        pct +
        "%;background:" +
        i.color +
        ';"></div></div><div class="val">' +
        fmtPct(pct) +
        "</div></div>";
    });
    document.getElementById(elId).innerHTML = html + "</div>";
  }

  function renderSOTable(ROWS_S) {
    var byOrder = {};
    ROWS_S.forEach(function (r) {
      if (r._covBucket !== "NOT COVERED" || !r.sales_order) return;
      if (!byOrder[r.sales_order])
        byOrder[r.sales_order] = {
          so: r.sales_order,
          lines: 0,
          fgSet: {},
          earliest: null,
          plantSet: {},
        };
      var o = byOrder[r.sales_order];
      o.lines++;
      if (r.fg_material) o.fgSet[r.fg_material] = 1;
      if (r.plant) o.plantSet[r.plant] = 1;
      if (
        r.requirement_date &&
        (!o.earliest || r.requirement_date < o.earliest)
      )
        o.earliest = r.requirement_date;
    });
    var arr = Object.keys(byOrder)
      .map(function (k) {
        return byOrder[k];
      })
      .sort(function (a, b) {
        return b.lines - a.lines;
      })
      .slice(0, 25);
    var rows = arr
      .map(function (o) {
        return (
          '<tr><td class="mono-code">' +
          esc(o.so) +
          '</td><td class="num">' +
          o.lines +
          "</td><td>" +
          Object.keys(o.fgSet).length +
          " FG</td><td>" +
          esc(Object.keys(o.plantSet).join(", ")) +
          "</td><td>" +
          fmtDate(o.earliest) +
          "</td></tr>"
        );
      })
      .join("");
    document.getElementById("soTable").innerHTML =
      "<thead><tr><th>Sales order</th><th>Short lines</th><th>FG materials</th><th>Plant(s)</th><th>Earliest requirement date</th></tr></thead><tbody>" +
      (rows ||
        '<tr><td colspan="5" class="empty-note">No uncovered sales orders in this view.</td></tr>') +
      "</tbody>";
  }

  /* ================= FG FOCUS ================= */
  var fgFocusState = { page: 0, pageSize: 25, selected: null };
  function fgFiltered() {
    var q = document.getElementById("fgSearch").value.trim().toLowerCase();
    var sort = document.getElementById("fgSort").value;
    var arr = FG_AGG;
    if (q) {
      arr = arr.filter(function (o) {
        return (
          (o.fg_material + " " + o.fg_description).toLowerCase().indexOf(q) !==
          -1
        );
      });
    }
    arr = arr.slice();
    if (sort === "short_desc")
      arr.sort(function (a, b) {
        return (
          b.componentsShort +
          b.componentsPartial -
          (a.componentsShort + a.componentsPartial)
        );
      });
    else if (sort === "qty_desc")
      arr.sort(function (a, b) {
        return b.totalOpenQty - a.totalOpenQty;
      });
    else if (sort === "readiness_asc")
      arr.sort(function (a, b) {
        return a.readiness - b.readiness;
      });
    else if (sort === "date_asc")
      arr.sort(function (a, b) {
        var av = a.earliest ? a.earliest.getTime() : 9e15,
          bv = b.earliest ? b.earliest.getTime() : 9e15;
        return av - bv;
      });
    return arr;
  }
  function renderFgList() {
    var filtered = fgFiltered();
    var totalPages = Math.max(
      1,
      Math.ceil(filtered.length / fgFocusState.pageSize),
    );
    if (fgFocusState.page >= totalPages) fgFocusState.page = totalPages - 1;
    var start = fgFocusState.page * fgFocusState.pageSize;
    var pageRows = filtered.slice(start, start + fgFocusState.pageSize);
    document.getElementById("fgListCount").textContent =
      fmtInt(filtered.length) + " matching FG materials";
    document.getElementById("fgPageInfo").textContent =
      "Page " + (fgFocusState.page + 1) + " of " + totalPages;
    document.getElementById("fgPrevPage").disabled = fgFocusState.page <= 0;
    document.getElementById("fgNextPage").disabled =
      fgFocusState.page >= totalPages - 1;
    var thead =
      "<thead><tr><th>FG material</th><th>Description</th><th>Plant(s)</th><th>Open qty</th><th>Components (short/total)</th><th>Readiness</th><th>Earliest due</th></tr></thead>";
    var tbody =
      "<tbody>" +
      pageRows
        .map(function (o) {
          var shortN = o.componentsShort + o.componentsPartial;
          var color =
            o.readiness >= 90
              ? "#2E8B57"
              : o.readiness >= 60
                ? "#C67C1E"
                : "#B0362A";
          return (
            '<tr class="clickable" data-fg="' +
            esc(o.fg_material) +
            '"><td class="mono-code">' +
            esc(o.fg_material) +
            "</td><td>" +
            esc(o.fg_description) +
            "</td>" +
            "<td>" +
            esc(o.plants.join(", ")) +
            '</td><td class="num">' +
            fmtNum(o.totalOpenQty) +
            "</td>" +
            '<td class="num">' +
            shortN +
            " / " +
            o.componentsTotal +
            "</td>" +
            '<td><div class="readiness"><div class="track"><div class="fill" style="width:' +
            o.readiness +
            "%;background:" +
            color +
            ';"></div></div>' +
            o.readiness +
            "%</div></td>" +
            "<td>" +
            fmtDate(o.earliest) +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody>";
    if (!pageRows.length)
      tbody =
        '<tbody><tr><td colspan="7" class="empty-note">No FG materials match this search.</td></tr></tbody>';
    document.getElementById("fgListTable").innerHTML = thead + tbody;
    document
      .querySelectorAll("#fgListTable tbody tr[data-fg]")
      .forEach(function (tr) {
        tr.addEventListener("click", function () {
          focusFg(tr.dataset.fg);
        });
      });
  }

  function focusFg(fgCode) {
    var o = FG_AGG_BY_CODE[fgCode];
    if (!o) {
      document.getElementById("fgDetailCard").style.display = "none";
      return;
    }
    fgFocusState.selected = fgCode;
    document.getElementById("fgDetailCard").style.display = "block";
    document.getElementById("fgDetailCode").textContent = fgCode;
    document.getElementById("fgDetailName").textContent =
      o.fg_description || "(no description)";
    var stats = [
      { l: "Plants", v: o.plants.join(", ") || "&mdash;" },
      { l: "Open demand qty", v: fmtNum(o.totalOpenQty) },
      { l: "Sales orders", v: fmtInt(o.soCount) },
      { l: "Components required", v: fmtInt(o.componentsTotal) },
      {
        l: "Blocked components",
        v: fmtInt(o.componentsShort + o.componentsPartial),
      },
      { l: "Readiness", v: o.readiness + "%" },
      { l: "Earliest due date", v: fmtDate(o.earliest) },
    ];
    document.getElementById("fgDetailStats").innerHTML = stats
      .map(function (s) {
        return (
          '<div class="fg-stat"><div class="l">' +
          s.l +
          '</div><div class="v">' +
          s.v +
          "</div></div>"
        );
      })
      .join("");

    var rows = rowsForPlant(LEVEL1_ROWS).filter(function (r) {
      return r.fg_material === fgCode;
    });
    var byComp = {};
    rows.forEach(function (r) {
      var k = r.component_material + "|" + r.plant;
      if (!byComp[k])
        byComp[k] = {
          component_material: r.component_material,
          component_description: r.component_description,
          uom: r.component_uom,
          plant: r.plant,
          totalShortfall: 0,
          lines: 0,
          worstBucket: "FULLY",
        };
      var c = byComp[k];
      c.totalShortfall += r.net_demand_this_line;
      c.lines++;
      if (r._covBucket === "NOT COVERED") c.worstBucket = "NOT COVERED";
      else if (r._covBucket === "PARTIALLY" && c.worstBucket !== "NOT COVERED")
        c.worstBucket = "PARTIALLY";
    });
    var compArr = Object.keys(byComp)
      .map(function (k) {
        return byComp[k];
      })
      .sort(function (a, b) {
        return b.totalShortfall - a.totalShortfall;
      });
    var thead =
      "<thead><tr><th>Component</th><th>Description</th><th>Plant</th><th>UOM</th><th>Available stock</th><th>Net shortfall qty</th><th>Lines</th><th>Status</th><th></th></tr></thead>";
    var tbody =
      "<tbody>" +
      compArr
        .map(function (c) {
          var stock = getStock(c.component_material, c.plant, 1);
          var simBtn =
            c.worstBucket !== "FULLY"
              ? '<button class="btn btn-outline btn-sm" data-sim-comp="' +
                esc(c.component_material) +
                '" data-sim-plant="' +
                esc(c.plant) +
                '" data-sim-level="1">Simulate</button>'
              : "";
          return (
            '<tr><td class="mono-code">' +
            esc(c.component_material) +
            "</td><td>" +
            esc(c.component_description) +
            "</td><td>" +
            esc(c.plant) +
            "</td><td>" +
            esc(c.uom) +
            "</td>" +
            '<td class="num">' +
            (stock === null ? "&mdash;" : fmtNum(stock)) +
            "</td>" +
            '<td class="num">' +
            fmtNum(c.totalShortfall) +
            '</td><td class="num">' +
            c.lines +
            "</td><td>" +
            pillHtml(c.worstBucket) +
            "</td><td>" +
            simBtn +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody>";
    document.getElementById("fgDetailTable").innerHTML = thead + tbody;
    document
      .querySelectorAll("#fgDetailTable [data-sim-comp]")
      .forEach(function (btn) {
        btn.addEventListener("click", function () {
          jumpToSimulator(
            btn.dataset.simComp,
            btn.dataset.simPlant,
            btn.dataset.simLevel,
          );
        });
      });
  }

  /* ================= PROCUREMENT PRIORITY ================= */
  var ppState = { page: 0, pageSize: 25 };
  function computeComponentAgg(scope, coverageMode, ROWS_S) {
    var srcRows =
      scope === "direct"
        ? ROWS_S.filter(function (r) {
            return r.bom_level === 1;
          })
        : ROWS_S;
    var map = {};
    srcRows.forEach(function (r) {
      if (coverageMode === "notcovered" && r._covBucket !== "NOT COVERED")
        return;
      if (coverageMode === "short" && !isShortBucket(r._covBucket)) return;
      var k = r.component_material;
      if (!k) return;
      if (!map[k])
        map[k] = {
          component_material: k,
          component_description: r.component_description,
          uom: r.component_uom,
          total: 0,
          lines: 0,
          pastDueLines: 0,
          fgSet: {},
          soSet: {},
          plantSet: {},
        };
      var o = map[k];
      o.total += r.net_demand_this_line;
      o.lines++;
      if (r._monthDiff !== null && r._monthDiff >= 1) o.pastDueLines++;
      if (r.fg_material) o.fgSet[r.fg_material] = 1;
      if (r.sales_order) o.soSet[r.sales_order] = 1;
      if (r.plant) o.plantSet[r.plant] = 1;
    });
    return Object.keys(map).map(function (k) {
      var o = map[k];
      o.fgSet_n = Object.keys(o.fgSet).length;
      o.soSet_n = Object.keys(o.soSet).length;
      o.plants = Object.keys(o.plantSet);
      return o;
    });
  }
  function renderPpTable() {
    var scope = document.getElementById("ppScope").value;
    var cov = document.getElementById("ppCoverage").value;
    var q = document.getElementById("ppSearch").value.trim().toLowerCase();
    var arr = computeComponentAgg(scope, cov, rowsForPlant(ROWS));
    if (q) {
      arr = arr.filter(function (o) {
        return (
          (o.component_material + " " + o.component_description)
            .toLowerCase()
            .indexOf(q) !== -1
        );
      });
    }
    arr.sort(function (a, b) {
      return b.fgSet_n - a.fgSet_n || b.total - a.total;
    });

    var totalPages = Math.max(1, Math.ceil(arr.length / ppState.pageSize));
    if (ppState.page >= totalPages) ppState.page = totalPages - 1;
    var start = ppState.page * ppState.pageSize;
    var pageRows = arr.slice(start, start + ppState.pageSize);
    document.getElementById("ppCount").textContent =
      fmtInt(arr.length) +
      " components" +
      (scope === "direct" ? " (direct Level-1 consumption)" : " (all levels)");
    document.getElementById("ppPageInfo").textContent =
      "Page " + (ppState.page + 1) + " of " + totalPages;
    document.getElementById("ppPrevPage").disabled = ppState.page <= 0;
    document.getElementById("ppNextPage").disabled =
      ppState.page >= totalPages - 1;

    var showStock = scope === "direct";
    var thead =
      "<thead><tr><th>Rank</th><th>Component</th><th>Description</th><th>UOM</th><th>Plant(s)</th>" +
      (showStock ? "<th>Stock avail.</th>" : "") +
      "<th>FGs impacted</th><th>Sales orders impacted</th><th>Net shortfall qty</th><th>Past-due lines</th><th></th></tr></thead>";
    var tbody =
      "<tbody>" +
      pageRows
        .map(function (o, idx) {
          var plantForSim = o.plants[0] || "";
          var stockCell = "";
          if (showStock) {
            var stockSum = 0,
              any = false;
            o.plants.forEach(function (p) {
              var s = getStock(o.component_material, p, 1);
              if (s !== null) {
                stockSum += s;
                any = true;
              }
            });
            stockCell =
              '<td class="num">' +
              (any ? fmtNum(stockSum) : "&mdash;") +
              "</td>";
          }
          return (
            '<tr><td class="num">' +
            (start + idx + 1) +
            '</td><td class="mono-code">' +
            esc(o.component_material) +
            "</td><td>" +
            esc(o.component_description) +
            "</td><td>" +
            esc(o.uom) +
            "</td>" +
            "<td>" +
            esc(o.plants.join(", ")) +
            "</td>" +
            stockCell +
            '<td class="num"><b>' +
            o.fgSet_n +
            '</b></td><td class="num">' +
            o.soSet_n +
            "</td>" +
            '<td class="num">' +
            fmtNum(o.total) +
            '</td><td class="num">' +
            o.pastDueLines +
            "</td>" +
            '<td><button class="btn btn-outline btn-sm" data-sim-comp="' +
            esc(o.component_material) +
            '" data-sim-plant="' +
            esc(plantForSim) +
            '">Simulate</button></td></tr>'
          );
        })
        .join("") +
      "</tbody>";
    if (!pageRows.length)
      tbody =
        '<tbody><tr><td colspan="10" class="empty-note">No components match this filter.</td></tr></tbody>';
    document.getElementById("ppTable").innerHTML = thead + tbody;
    document
      .querySelectorAll("#ppTable [data-sim-comp]")
      .forEach(function (btn) {
        btn.addEventListener("click", function () {
          jumpToSimulator(btn.dataset.simComp, btn.dataset.simPlant, null);
        });
      });
  }

  /* ================= STOCK VISIBILITY ================= */
  var svState = { page: 0, pageSize: 30 };
  function renderStockVis() {
    var levelSel = document.getElementById("svLevel");
    if (!levelSel.options.length) {
      var levels = {};
      ROWS.forEach(function (r) {
        levels[r.bom_level] = 1;
      });
      var arr = Object.keys(levels)
        .map(Number)
        .sort(function (a, b) {
          return a - b;
        });
      levelSel.innerHTML = arr
        .map(function (l) {
          return (
            '<option value="' +
            l +
            '"' +
            (l === 1 ? " selected" : "") +
            ">Level " +
            l +
            (l === 1 ? " (direct FG consumption)" : "") +
            "</option>"
          );
        })
        .join("");
    }
    var level = levelSel.value || "1";
    var show = document.getElementById("svShow").value;
    var q = document.getElementById("svSearch").value.trim().toLowerCase();
    var keys = Object.keys(STOCK_INDEX).filter(function (k) {
      return STOCK_INDEX[k].level === Number(level);
    });
    var arr = keys.map(function (k) {
      return STOCK_INDEX[k];
    });
    if (show === "short")
      arr = arr.filter(function (o) {
        return o.netShortfall > 0.000001;
      });
    if (q)
      arr = arr.filter(function (o) {
        return (
          (o.component_material + " " + o.component_description)
            .toLowerCase()
            .indexOf(q) !== -1
        );
      });
    arr.sort(function (a, b) {
      return b.netShortfall - a.netShortfall;
    });

    var totalPages = Math.max(1, Math.ceil(arr.length / svState.pageSize));
    if (svState.page >= totalPages) svState.page = totalPages - 1;
    var start = svState.page * svState.pageSize;
    var pageRows = arr.slice(start, start + svState.pageSize);
    document.getElementById("svCount").textContent =
      fmtInt(arr.length) + " component/plant combinations at Level " + level;
    document.getElementById("svPageInfo").textContent =
      "Page " + (svState.page + 1) + " of " + totalPages;
    document.getElementById("svPrevPage").disabled = svState.page <= 0;
    document.getElementById("svNextPage").disabled =
      svState.page >= totalPages - 1;

    var thead =
      "<thead><tr><th>Component</th><th>Description</th><th>Plant</th><th>UOM</th><th>Available stock</th><th>Total gross demand</th><th>Net shortfall</th><th>Coverage</th><th>FGs relying on it</th></tr></thead>";
    var tbody =
      "<tbody>" +
      pageRows
        .map(function (o) {
          var covPct =
            o.grossDemand > 0
              ? Math.max(
                  0,
                  Math.min(100, Math.round((o.stock / o.grossDemand) * 100)),
                )
              : 100;
          var color =
            covPct >= 100 ? "#2E8B57" : covPct >= 50 ? "#C67C1E" : "#B0362A";
          return (
            '<tr><td class="mono-code">' +
            esc(o.component_material) +
            "</td><td>" +
            esc(o.component_description) +
            "</td><td>" +
            esc(o.plant) +
            "</td><td>" +
            esc(o.uom) +
            "</td>" +
            '<td class="num">' +
            fmtNum(o.stock) +
            '</td><td class="num">' +
            fmtNum(o.grossDemand) +
            '</td><td class="num">' +
            fmtNum(o.netShortfall) +
            "</td>" +
            '<td><div class="readiness"><div class="track"><div class="fill" style="width:' +
            covPct +
            "%;background:" +
            color +
            ';"></div></div>' +
            covPct +
            "%</div></td>" +
            '<td class="num">' +
            o.fgCount +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody>";
    if (!pageRows.length)
      tbody =
        '<tbody><tr><td colspan="9" class="empty-note">No components match this filter.</td></tr></tbody>';
    document.getElementById("svTable").innerHTML = thead + tbody;
  }

  /* ================= SUPPLY SIMULATOR ================= */
  function buildShortIndex(ROWS_S) {
    SHORT_ROWS_BY_COMP = {};
    ROWS_S.forEach(function (r) {
      if (!isShortBucket(r._covBucket)) return;
      if (!SHORT_ROWS_BY_COMP[r.component_material])
        SHORT_ROWS_BY_COMP[r.component_material] = [];
      SHORT_ROWS_BY_COMP[r.component_material].push(r);
    });
  }
  function refreshSimCompOptions() {
    var q = document.getElementById("simCompSearch").value.trim().toLowerCase();
    var codes = Object.keys(SHORT_ROWS_BY_COMP);
    var descByCode = {};
    codes.forEach(function (c) {
      descByCode[c] = SHORT_ROWS_BY_COMP[c][0].component_description;
    });
    var matches = codes.filter(function (c) {
      return !q || (c + " " + descByCode[c]).toLowerCase().indexOf(q) !== -1;
    });
    matches.sort(function (a, b) {
      return SHORT_ROWS_BY_COMP[b].length - SHORT_ROWS_BY_COMP[a].length;
    });
    matches = matches.slice(0, 200);
    var sel = document.getElementById("simCompSelect");
    sel.innerHTML = matches
      .map(function (c) {
        return (
          '<option value="' +
          esc(c) +
          '">' +
          esc(c) +
          " \u2014 " +
          esc(descByCode[c] || "") +
          "</option>"
        );
      })
      .join("");
    sel.size = Math.min(8, Math.max(3, matches.length || 1));
    document.getElementById("simResult").classList.remove("show");
    document.getElementById("simStockLine").textContent = "";
    if (matches.length) {
      sel.selectedIndex = 0;
      refreshSimPlantLevelOptions();
      document.getElementById("simEmpty").style.display = "none";
    } else {
      document.getElementById("simPlantSelect").innerHTML = "";
      document.getElementById("simLevelSelect").innerHTML = "";
      document.getElementById("simEmpty").style.display = "block";
      document.getElementById("simEmpty").textContent =
        "No components with open shortfall match this search (in the current plant view).";
    }
  }
  function refreshSimPlantLevelOptions() {
    var comp = document.getElementById("simCompSelect").value;
    var rows = SHORT_ROWS_BY_COMP[comp] || [];
    var plants = {};
    rows.forEach(function (r) {
      plants[r.plant] = (plants[r.plant] || 0) + 1;
    });
    var plantArr = Object.keys(plants).sort(function (a, b) {
      return plants[b] - plants[a];
    });
    var pSel = document.getElementById("simPlantSelect");
    pSel.innerHTML = plantArr
      .map(function (p) {
        return (
          '<option value="' +
          esc(p) +
          '">' +
          esc(p) +
          " (" +
          plants[p] +
          " short lines)</option>"
        );
      })
      .join("");
    pSel.onchange = refreshSimLevelOptions;
    refreshSimLevelOptions();
  }
  function refreshSimLevelOptions() {
    var comp = document.getElementById("simCompSelect").value;
    var plant = document.getElementById("simPlantSelect").value;
    var rows = (SHORT_ROWS_BY_COMP[comp] || []).filter(function (r) {
      return r.plant === plant;
    });
    var levels = {};
    rows.forEach(function (r) {
      levels[r.bom_level] = (levels[r.bom_level] || 0) + r.net_demand_this_line;
    });
    var levelArr = Object.keys(levels)
      .map(Number)
      .sort(function (a, b) {
        return levels[b] - levels[a];
      });
    var lSel = document.getElementById("simLevelSelect");
    lSel.innerHTML = levelArr
      .map(function (l) {
        return (
          '<option value="' +
          l +
          '">Level ' +
          l +
          " (" +
          fmtNum(levels[l]) +
          " units short)</option>"
        );
      })
      .join("");
    updateSimStockLine();
    lSel.onchange = updateSimStockLine;
  }
  function updateSimStockLine() {
    var comp = document.getElementById("simCompSelect").value;
    var plant = document.getElementById("simPlantSelect").value;
    var level = document.getElementById("simLevelSelect").value;
    if (!comp || !plant || !level) {
      document.getElementById("simStockLine").textContent = "";
      return;
    }
    var stock = getStock(comp, plant, level);
    var idxRec = STOCK_INDEX[comp + "|" + plant + "|" + level];
    if (stock === null || !idxRec) {
      document.getElementById("simStockLine").textContent = "";
      return;
    }
    document.getElementById("simStockLine").innerHTML =
      "Current stock on hand: <b>" +
      fmtNum(stock) +
      "</b> \u00b7 total open gross demand at this level: <b>" +
      fmtNum(idxRec.grossDemand) +
      "</b> \u00b7 total net shortfall: <b>" +
      fmtNum(idxRec.netShortfall) +
      "</b>";
  }
  function jumpToSimulator(comp, plant, level) {
    activateTab("simulator");
    document.getElementById("simCompSearch").value = comp;
    refreshSimCompOptions();
    var sel = document.getElementById("simCompSelect");
    for (var i = 0; i < sel.options.length; i++) {
      if (sel.options[i].value === comp) {
        sel.selectedIndex = i;
        break;
      }
    }
    refreshSimPlantLevelOptions();
    if (plant) {
      var pSel = document.getElementById("simPlantSelect");
      for (var j = 0; j < pSel.options.length; j++) {
        if (pSel.options[j].value === plant) {
          pSel.selectedIndex = j;
          break;
        }
      }
      refreshSimLevelOptions();
    }
    if (level) {
      var lSel = document.getElementById("simLevelSelect");
      for (var k = 0; k < lSel.options.length; k++) {
        if (lSel.options[k].value === String(level)) {
          lSel.selectedIndex = k;
          break;
        }
      }
      updateSimStockLine();
    }
    document.getElementById("simQty").focus();
  }
  function runSimulation() {
    var comp = document.getElementById("simCompSelect").value;
    var plant = document.getElementById("simPlantSelect").value;
    var level = document.getElementById("simLevelSelect").value;
    var qty = Number(document.getElementById("simQty").value) || 0;
    if (!comp || !plant || !level) {
      return;
    }
    var rows = ROWS.filter(function (r) {
      return (
        r.component_material === comp &&
        r.plant === plant &&
        String(r.bom_level) === String(level) &&
        r.net_demand_this_line > 0
      );
    });
    rows.sort(function (a, b) {
      return a._sortKey < b._sortKey ? -1 : a._sortKey > b._sortKey ? 1 : 0;
    });

    var running = 0,
      cleared = [],
      remaining = [];
    rows.forEach(function (r) {
      var before = running;
      running += r.net_demand_this_line;
      if (running <= qty + 1e-6) {
        cleared.push(r);
      } else if (before < qty) {
        var copy = {};
        for (var p in r) copy[p] = r[p];
        copy._partial = true;
        copy._clearedQty = Math.max(0, qty - before);
        cleared.push(copy);
      } else {
        remaining.push(r);
      }
    });
    var clearedSOs = {},
      clearedFGs = {};
    cleared.forEach(function (r) {
      if (r.sales_order) clearedSOs[r.sales_order] = 1;
      if (r.fg_material) clearedFGs[r.fg_material] = 1;
    });
    var remainingSOs = {},
      remainingFGs = {};
    remaining.forEach(function (r) {
      if (r.sales_order) remainingSOs[r.sales_order] = 1;
      if (r.fg_material) remainingFGs[r.fg_material] = 1;
    });
    var totalShort = rows.reduce(function (s, r) {
      return s + r.net_demand_this_line;
    }, 0);
    var leftoverQty = Math.max(0, qty - totalShort);

    document.getElementById("simResult").classList.add("show");

    var fullyClearedCount = cleared.filter(function (r) {
      return !r._partial;
    }).length;
    document.getElementById("simHeadline").textContent =
      fmtNum(Math.min(qty, totalShort)) +
      " of " +
      fmtNum(totalShort) +
      " units of open shortfall covered";
    document.getElementById("simSubline").textContent =
      fullyClearedCount +
      " of " +
      rows.length +
      " open lines fully clear \u2192 " +
      Object.keys(clearedSOs).length +
      " sales order(s) and " +
      Object.keys(clearedFGs).length +
      " FG material(s) get this component fully satisfied." +
      (leftoverQty > 0
        ? " " +
          fmtNum(leftoverQty) +
          " units left over beyond all currently open shortfall at this level."
        : "");

    document.getElementById("simClearedSub").textContent =
      fullyClearedCount +
      " lines clear completely, in priority order (earliest requirement date first)";
    var ct =
      "<thead><tr><th>FG material</th><th>Sales order</th><th>Req. date</th><th>Qty needed</th><th></th></tr></thead><tbody>" +
      cleared
        .slice(0, 50)
        .map(function (r) {
          return (
            '<tr><td class="mono-code">' +
            esc(r.fg_material) +
            "</td><td>" +
            esc(r.sales_order) +
            "</td><td>" +
            fmtDate(r.requirement_date) +
            "</td>" +
            '<td class="num">' +
            fmtNum(r._partial ? r._clearedQty : r.net_demand_this_line) +
            "</td>" +
            "<td>" +
            (r._partial
              ? '<span class="pill amber">partial</span>'
              : '<span class="pill green">clears</span>') +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody>";
    document.getElementById("simClearedTable").innerHTML = cleared.length
      ? ct
      : '<tbody><tr><td class="empty-note">No lines clear at this quantity.</td></tr></tbody>';

    document.getElementById("simRemainSub").textContent =
      remaining.length + " lines still short after this injection";
    var rt =
      "<thead><tr><th>FG material</th><th>Sales order</th><th>Req. date</th><th>Still short</th></tr></thead><tbody>" +
      remaining
        .slice(0, 50)
        .map(function (r) {
          return (
            '<tr><td class="mono-code">' +
            esc(r.fg_material) +
            "</td><td>" +
            esc(r.sales_order) +
            "</td><td>" +
            fmtDate(r.requirement_date) +
            '</td><td class="num">' +
            fmtNum(r.net_demand_this_line) +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody>";
    document.getElementById("simRemainTable").innerHTML = remaining.length
      ? rt
      : '<tbody><tr><td class="empty-note">Nothing left short \u2014 this quantity fully clears the queue at this level.</td></tr></tbody>';
  }

  /* ================= AGING / TIME RISK (month-based) ================= */
  function bucketForMonth(md) {
    if (md === null) return "No date";
    if (md >= 3) return "Old demand (3+ months) \u2014 verify";
    if (md >= 1) return "Past due (1\u20132 months)";
    if (md === 0) return "Current month";
    if (md === -1) return "Next month";
    if (md === -2) return "2 months out";
    return "3+ months out";
  }
  var BUCKET_ORDER = [
    "Old demand (3+ months) \u2014 verify",
    "Past due (1\u20132 months)",
    "Current month",
    "Next month",
    "2 months out",
    "3+ months out",
    "No date",
  ];
  var BUCKET_COLORS = {
    "Old demand (3+ months) \u2014 verify": "#6A4C93",
    "Past due (1\u20132 months)": "#B0362A",
    "Current month": "#1C7293",
    "Next month": "#C67C1E",
    "2 months out": "#9AA7B4",
    "3+ months out": "#C7CFD6",
    "No date": "#DAE1E8",
  };
  function renderTimeChart(ROWS_S) {
    var counts = {};
    BUCKET_ORDER.forEach(function (b) {
      counts[b] = 0;
    });
    ROWS_S.forEach(function (r) {
      if (isShortBucket(r._covBucket)) {
        counts[bucketForMonth(r._monthDiff)]++;
      }
    });
    var items = BUCKET_ORDER.map(function (b) {
      return { label: b, value: counts[b], color: BUCKET_COLORS[b] };
    });
    renderBarList("timeChart", items);
  }
  function renderPastDueTable(ROWS_S) {
    var byOrder = {};
    ROWS_S.forEach(function (r) {
      if (
        !(r._covBucket === "NOT COVERED") ||
        r._monthDiff === null ||
        r._monthDiff < 1 ||
        !r.sales_order
      )
        return;
      if (!byOrder[r.sales_order])
        byOrder[r.sales_order] = {
          so: r.sales_order,
          lines: 0,
          maxMonthsLate: 0,
          fgSet: {},
        };
      var o = byOrder[r.sales_order];
      o.lines++;
      o.maxMonthsLate = Math.max(o.maxMonthsLate, r._monthDiff);
      if (r.fg_material) o.fgSet[r.fg_material] = 1;
    });
    var arr = Object.keys(byOrder)
      .map(function (k) {
        return byOrder[k];
      })
      .sort(function (a, b) {
        return b.maxMonthsLate - a.maxMonthsLate;
      })
      .slice(0, 25);
    var rows = arr
      .map(function (o) {
        var tag =
          o.maxMonthsLate >= 3
            ? '<span class="pill purple">' +
              o.maxMonthsLate +
              " mo \u2014 verify</span>"
            : '<span class="pill red">' + o.maxMonthsLate + " mo late</span>";
        return (
          '<tr><td class="mono-code">' +
          esc(o.so) +
          '</td><td class="num">' +
          o.lines +
          "</td><td>" +
          tag +
          "</td><td>" +
          Object.keys(o.fgSet).length +
          " FG</td></tr>"
        );
      })
      .join("");
    document.getElementById("pastDueTable").innerHTML =
      "<thead><tr><th>Sales order</th><th>Short lines</th><th>Age</th><th>FG materials</th></tr></thead><tbody>" +
      (rows ||
        '<tr><td colspan="4" class="empty-note">No aged uncovered demand found.</td></tr>') +
      "</tbody>";
  }

  /* ================= PLANTS (always global) ================= */
  function renderPlantTable() {
    var map = {};
    ROWS.forEach(function (r) {
      var p = r.plant || "(blank)";
      if (!map[p])
        map[p] = {
          plant: p,
          total: 0,
          notCovered: 0,
          partial: 0,
          soSet: {},
          fgSet: {},
          notCoveredSoSet: {},
        };
      var o = map[p];
      o.total++;
      if (r.sales_order) o.soSet[r.sales_order] = 1;
      if (r.fg_material) o.fgSet[r.fg_material] = 1;
      if (r._covBucket === "NOT COVERED") {
        o.notCovered++;
        if (r.sales_order) o.notCoveredSoSet[r.sales_order] = 1;
      }
      if (r._covBucket === "PARTIALLY") o.partial++;
    });
    var arr = Object.keys(map)
      .map(function (k) {
        return map[k];
      })
      .sort(function (a, b) {
        return b.notCovered - a.notCovered;
      });
    var rows = arr
      .map(function (o) {
        var pct = o.total ? (o.notCovered / o.total) * 100 : 0;
        var isCur = o.plant === CURRENT_PLANT;
        return (
          "<tr" +
          (isCur ? ' style="background:var(--teal-light);"' : "") +
          '><td class="mono-code">' +
          esc(o.plant) +
          (isCur ? ' <span class="pill green">viewing</span>' : "") +
          '</td><td class="num">' +
          fmtInt(o.total) +
          '</td><td class="num">' +
          fmtInt(o.notCovered) +
          '</td><td class="num">' +
          fmtInt(o.partial) +
          '</td><td class="num">' +
          fmtPct(pct) +
          '</td><td class="num">' +
          Object.keys(o.notCoveredSoSet).length +
          '</td><td class="num">' +
          Object.keys(o.fgSet).length +
          "</td></tr>"
        );
      })
      .join("");
    document.getElementById("plantTable").innerHTML =
      "<thead><tr><th>Plant</th><th>Total lines</th><th>Not covered</th><th>Partially covered</th><th>% not covered</th><th>SOs affected</th><th>FGs in scope</th></tr></thead><tbody>" +
      (rows || '<tr><td colspan="7" class="empty-note">No data.</td></tr>') +
      "</tbody>";
  }

  /* ================= INTEGRITY (always global) ================= */
  function renderIntegrity() {
    var seen = {},
      dupCount = 0,
      netGreaterThanGross = 0,
      negativeNet = 0,
      mismatch = 0;
    ROWS.forEach(function (r) {
      seen[r._key] = (seen[r._key] || 0) + 1;
      if (r.net_demand_this_line > r.gross_child_part_demand_to_cover_fg + 0.01)
        netGreaterThanGross++;
      if (r.net_demand_this_line < -0.01) negativeNet++;
      var expected;
      if (r.net_demand_this_line <= 0.000001) expected = "FULLY";
      else if (
        r.net_demand_this_line <
        r.gross_child_part_demand_to_cover_fg - 0.01
      )
        expected = "PARTIALLY";
      else expected = "NOT COVERED";
      if (expected !== r._covBucket) mismatch++;
    });
    Object.keys(seen).forEach(function (k) {
      if (seen[k] > 1) dupCount += seen[k];
    });
    var checks = [
      {
        t: "Duplicate key rows",
        n: dupCount,
        good: dupCount === 0,
        detail:
          dupCount === 0
            ? "No duplicate (plant, component, SO, item, schedule line, BOM path) combinations."
            : fmtInt(dupCount) +
              " rows share a key \u2014 re-check the stpo_deduplicated / stko_deduplicated logic for this scope.",
      },
      {
        t: "Net exceeds gross",
        n: netGreaterThanGross,
        good: netGreaterThanGross === 0,
        detail:
          netGreaterThanGross === 0
            ? "Net demand never exceeds gross demand on any line."
            : fmtInt(netGreaterThanGross) +
              " rows show net > gross \u2014 investigate the netting CTEs.",
      },
      {
        t: "Negative net shortfall",
        n: negativeNet,
        good: negativeNet === 0,
        detail:
          negativeNet === 0
            ? "No negative net_demand_this_line values."
            : fmtInt(negativeNet) + " rows have negative net demand.",
      },
      {
        t: "Coverage flag mismatches",
        n: mismatch,
        good: mismatch === 0,
        detail:
          mismatch === 0
            ? "net_coverage_flag agrees with net_demand_this_line vs gross on every row."
            : fmtInt(mismatch) + " rows have an inconsistent coverage flag.",
      },
    ];
    document.getElementById("integrityList").innerHTML = checks
      .map(function (c) {
        return (
          '<div class="integrity-item"><div><div class="t">' +
          c.t +
          '</div><div class="n">' +
          esc(c.detail) +
          '</div></div><div class="check-badge">' +
          (c.good
            ? '<span style="color:#2E8B57;">&#10003;</span>'
            : '<span style="color:#B0362A;">&#10007; ' +
              fmtInt(c.n) +
              "</span>") +
          "</div></div>"
        );
      })
      .join("");
  }
  function renderScopeDetail() {
    var plantSet = {};
    ROWS.forEach(function (r) {
      if (r.plant) plantSet[r.plant] = 1;
    });
    var plants = Object.keys(plantSet).sort();
    var earliestDate = null,
      latestDate = null;
    ROWS.forEach(function (r) {
      if (r.requirement_date) {
        if (!earliestDate || r.requirement_date < earliestDate)
          earliestDate = r.requirement_date;
        if (!latestDate || r.requirement_date > latestDate)
          latestDate = r.requirement_date;
      }
    });
    var html =
      "<div><b>Plants in file (" +
      plants.length +
      "):</b> " +
      esc(plants.join(", ")) +
      "</div>";
    html +=
      "<div><b>Requirement date span:</b> " +
      fmtDate(earliestDate) +
      " to " +
      fmtDate(latestDate) +
      "</div>";
    html +=
      "<div><b>Analysis run as of:</b> " +
      fmtDate(AS_OF) +
      " (your device\u2019s current date)</div>";
    document.getElementById("scopeDetail").innerHTML = html;
  }

  /* ================= EXPLORER ================= */
  var explorerState = {
    page: 0,
    pageSize: 50,
    sortCol: "requirement_date",
    sortDir: 1,
  };
  function setupExplorerOptions() {
    var levelSet = {};
    ROWS.forEach(function (r) {
      levelSet[r.bom_level] = 1;
    });
    var levelKeys = Object.keys(levelSet)
      .map(Number)
      .sort(function (a, b) {
        return a - b;
      });
    document.getElementById("fLevel").innerHTML =
      '<option value="">All</option>' +
      levelKeys
        .map(function (l) {
          return '<option value="' + l + '">Level ' + l + "</option>";
        })
        .join("");
    var fp = document.getElementById("fPlant");
    fp.innerHTML =
      '<option value="">All</option>' +
      ALL_PLANTS.map(function (p) {
        return '<option value="' + esc(p) + '">' + esc(p) + "</option>";
      }).join("");
    fp.value = CURRENT_PLANT || "";
  }
  function explorerFiltered() {
    var plant = document.getElementById("fPlant").value;
    var level = document.getElementById("fLevel").value;
    var cov = document.getElementById("fCoverage").value;
    var minQty = Number(document.getElementById("fMinQty").value) || 0;
    var q = document.getElementById("fSearch").value.trim().toLowerCase();
    return ROWS.filter(function (r) {
      if (plant && r.plant !== plant) return false;
      if (level && String(r.bom_level) !== String(level)) return false;
      if (cov && r._covBucket !== cov) return false;
      if (minQty > 0 && r.net_demand_this_line < minQty) return false;
      if (q) {
        var hay = (
          r.fg_material +
          " " +
          r.fg_description +
          " " +
          r.component_material +
          " " +
          r.component_description +
          " " +
          r.sales_order
        ).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }
  var EXPLORER_COLS = [
    { k: "plant", label: "Plant" },
    { k: "fg_material", label: "FG material" },
    { k: "component_material", label: "Component" },
    { k: "bom_level", label: "Level", num: true },
    { k: "sales_order", label: "Sales order" },
    { k: "requirement_date", label: "Req. date", date: true },
    { k: "gross_child_part_stock", label: "Stock (running)", num: true },
    {
      k: "gross_child_part_demand_to_cover_fg",
      label: "Gross demand",
      num: true,
    },
    { k: "net_demand_this_line", label: "Net demand", num: true },
    { k: "_covBucket", label: "Coverage" },
  ];
  function renderExplorer() {
    var filtered = explorerFiltered();
    filtered.sort(function (a, b) {
      var c = explorerState.sortCol,
        av = a[c],
        bv = b[c];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (av < bv) return -1 * explorerState.sortDir;
      if (av > bv) return 1 * explorerState.sortDir;
      return 0;
    });
    var total = filtered.length;
    var totalPages = Math.max(1, Math.ceil(total / explorerState.pageSize));
    if (explorerState.page >= totalPages) explorerState.page = totalPages - 1;
    var start = explorerState.page * explorerState.pageSize;
    var pageRows = filtered.slice(start, start + explorerState.pageSize);
    document.getElementById("explorerCount").textContent =
      fmtInt(total) + " matching rows";
    document.getElementById("pageInfo").textContent =
      "Page " + (explorerState.page + 1) + " of " + totalPages;
    document.getElementById("prevPage").disabled = explorerState.page <= 0;
    document.getElementById("nextPage").disabled =
      explorerState.page >= totalPages - 1;
    var thead =
      "<thead><tr>" +
      EXPLORER_COLS.map(function (c) {
        var cls =
          c.k === explorerState.sortCol
            ? "sorted " + (explorerState.sortDir === 1 ? "asc" : "")
            : "";
        return (
          '<th data-col="' + c.k + '" class="' + cls + '">' + c.label + "</th>"
        );
      }).join("") +
      "</tr></thead>";
    var tbody =
      "<tbody>" +
      pageRows
        .map(function (r) {
          return (
            "<tr>" +
            EXPLORER_COLS.map(function (c) {
              var v = r[c.k];
              if (c.date) v = fmtDate(v);
              if (c.num)
                return (
                  '<td class="num">' +
                  (c.k === "bom_level" ? v : fmtNum(v)) +
                  "</td>"
                );
              if (c.k === "_covBucket") return "<td>" + pillHtml(v) + "</td>";
              if (c.k === "fg_material" || c.k === "component_material")
                return '<td class="mono-code">' + esc(v) + "</td>";
              return "<td>" + esc(v) + "</td>";
            }).join("") +
            "</tr>"
          );
        })
        .join("") +
      "</tbody>";
    if (!pageRows.length)
      tbody =
        '<tbody><tr><td colspan="' +
        EXPLORER_COLS.length +
        '" class="empty-note">No rows match these filters.</td></tr></tbody>';
    document.getElementById("explorerTable").innerHTML = thead + tbody;
    document.querySelectorAll("#explorerTable thead th").forEach(function (th) {
      th.addEventListener("click", function () {
        var col = th.dataset.col;
        if (explorerState.sortCol === col) {
          explorerState.sortDir *= -1;
        } else {
          explorerState.sortCol = col;
          explorerState.sortDir = 1;
        }
        renderExplorer();
      });
    });
  }
  function exportCSV() {
    var filtered = explorerFiltered();
    var cols = [
      "plant",
      "company_code",
      "fg_material",
      "fg_description",
      "bom_level",
      "parent_material",
      "bom_item_number",
      "component_material",
      "component_description",
      "sales_order",
      "item_number",
      "schedule_line",
      "requirement_date",
      "fg_demand",
      "bom_qty_per_fg",
      "gross_child_part_demand_to_cover_fg",
      "gross_child_part_stock",
      "gross_shortage_or_excess",
      "gross_status_flag",
      "net_demand_this_line",
      "net_coverage_flag",
      "component_uom",
    ];
    var lines = [cols.join(",")];
    filtered.forEach(function (r) {
      lines.push(
        cols
          .map(function (c) {
            var v = r[c];
            if (v instanceof Date) v = v.toISOString().slice(0, 10);
            v = v == null ? "" : String(v);
            if (v.indexOf(",") !== -1 || v.indexOf('"') !== -1)
              v = '"' + v.replace(/"/g, '""') + '"';
            return v;
          })
          .join(","),
      );
    });
    var csv = lines.join("\n");
    var filename =
      "control_tower_export_" + new Date().toISOString().slice(0, 10) + ".csv";
    if (downloadsCap) {
      downloadsCap.save({ filename: filename, data: csv }).catch(function () {
        fallbackDownload(csv, filename);
      });
    } else {
      fallbackDownload(csv, filename);
    }
  }
  function fallbackDownload(csv, filename) {
    try {
      var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function () {
        URL.revokeObjectURL(url);
      }, 1000);
    } catch (e) {
      alert("Export is not available in this view.");
    }
  }

  (function initCapabilities() {
    try {
      if (window.claude && typeof window.claude.use === "function") {
        window.claude
          .use("downloads")
          .then(function (cap) {
            downloadsCap = cap;
          })
          .catch(function () {});
      }
    } catch (e) {}
  })();
})();
