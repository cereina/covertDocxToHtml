#!/usr/bin/env node
/**
 * DOCX → HTML converter (GCWeb/WET-BOEW)
 *
 * Supports:
 * - details/summary using Word styles:
 *   - WET Details Summary
 *   - WET Details Content
 * - accordion using Word styles:
 *   - WET Accordion Start
 *   - WET Accordion Heading
 *   - WET Accordion Panel
 *   - WET Accordion End
 * - pagination using Word styles:
 *   - WET Pagination Start
 *   - WET Pagination Item
 *   - WET Pagination Active
 *   - WET Pagination Disabled
 *   - WET Pagination End
 *
 * Also maps common styles to classes (alerts, buttons, lead, etc.)
 *
 * Usage:
 *   node docx_to_gcweb_html_extended.js input.docx -o output.html
 */

const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const cheerio = require("cheerio");

function parseArgs(argv) {
  const args = { input: null, output: null };
  const positional = [];
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === "-o" || a === "--output") {
      args.output = argv[i + 1];
      i++;
    } else {
      positional.push(a);
    }
  }
  args.input = positional[0] || null;
  return args;
}

// Word style → placeholder HTML mapping (phase 1)
// We convert style names into recognizable marker elements (divs with data-wet="...")
// Then we transform those markers into proper structures (phase 2).
const styleMap = [
  // Typography / utilities
  "p[style-name='WET Lead'] => p.lead:fresh",
  "p[style-name='WET Small'] => p.small:fresh",
  "p[style-name='WET Muted'] => p.text-muted:fresh",
  "p[style-name='WET Blockquote'] => blockquote:fresh",

  // Alerts (wrap later by converting these paragraphs into <section class="alert ...">)
  "p[style-name='WET Alert Success'] => div[data-wet='alert-success'] > p:fresh",
  "p[style-name='WET Alert Info'] => div[data-wet='alert-info'] > p:fresh",
  "p[style-name='WET Alert Warning'] => div[data-wet='alert-warning'] > p:fresh",
  "p[style-name='WET Alert Danger'] => div[data-wet='alert-danger'] > p:fresh",

  // Well
  "p[style-name='WET Well'] => div.well > p:fresh",

  // Buttons (placeholder links)
  "p[style-name='WET Button Primary'] => a.btn.btn-primary[data-wet='btn'][href='#'][role='button']:fresh",
  "p[style-name='WET Button Default'] => a.btn.btn-default[data-wet='btn'][href='#'][role='button']:fresh",
  "p[style-name='WET Button Danger'] => a.btn.btn-danger[data-wet='btn'][href='#'][role='button']:fresh",
  "p[style-name='WET Button Link'] => a.btn.btn-link[data-wet='btn'][href='#'][role='button']:fresh",

  // Table markers (apply to next table in post-processing)
  "p[style-name='WET Table Basic'] => div[data-wet='table-basic']:fresh",
  "p[style-name='WET Table Striped'] => div[data-wet='table-striped']:fresh",
  "p[style-name='WET Table Bordered'] => div[data-wet='table-bordered']:fresh",
  "p[style-name='WET Table Hover'] => div[data-wet='table-hover']:fresh",
  "p[style-name='WET Table Condensed'] => div[data-wet='table-condensed']:fresh",
  "p[style-name='WET Table Responsive'] => div[data-wet='table-responsive']:fresh",

  // Details markers
  "p[style-name='WET Details Summary'] => div[data-wet='details-summary']:fresh",
  "p[style-name='WET Details Content'] => div[data-wet='details-content']:fresh",

  // Accordion markers
  "p[style-name='WET Accordion Start'] => div[data-wet='accordion-start']:fresh",
  "p[style-name='WET Accordion Heading'] => div[data-wet='accordion-heading']:fresh",
  "p[style-name='WET Accordion Panel'] => div[data-wet='accordion-panel']:fresh",
  "p[style-name='WET Accordion End'] => div[data-wet='accordion-end']:fresh",

  // Pagination markers
  "p[style-name='WET Pagination Start'] => div[data-wet='pagination-start']:fresh",
  "p[style-name='WET Pagination Item'] => div[data-wet='pagination-item']:fresh",
  "p[style-name='WET Pagination Active'] => div[data-wet='pagination-active']:fresh",
  "p[style-name='WET Pagination Disabled'] => div[data-wet='pagination-disabled']:fresh",
  "p[style-name='WET Pagination End'] => div[data-wet='pagination-end']:fresh",

    // Alignment utilities
  "p[style-name='WET Text Center'] => p.text-center:fresh",
  "p[style-name='WET Pull Right'] => p.pull-right:fresh",

  // List markers (we'll convert these marker divs to classes on the next <ul>/<ol>)
  "p[style-name='WET List Inline'] => div[data-wet='list-inline']:fresh",
  "p[style-name='WET List Unstyled'] => div[data-wet='list-unstyled']:fresh",

  // Badges (Bootstrap 3 style used in WET; keep simple: <span class='badge ...'>)
  "p[style-name='WET Badge Default'] => span.badge[data-wet='badge-default']:fresh",
  "p[style-name='WET Badge Primary'] => span.badge.badge-primary[data-wet='badge-primary']:fresh",
  "p[style-name='WET Badge Success'] => span.badge.badge-success[data-wet='badge-success']:fresh",
  "p[style-name='WET Badge Warning'] => span.badge.badge-warning[data-wet='badge-warning']:fresh",
  "p[style-name='WET Badge Danger'] => span.badge.badge-danger[data-wet='badge-danger']:fresh",

];

function transformPlaceholdersToGCWeb(html) {
  const $ = cheerio.load(html, { decodeEntities: false });

  // Wrap output in <main ... class="container">
  // If mammoth returns multiple roots, keep them inside main.
  const bodyChildren = $("body").length ? $("body").children().toArray() : $.root().children().toArray();
  const $main = $("<main/>")
    .attr("property", "mainContentOfPage")
    .addClass("container");

  bodyChildren.forEach((el) => $main.append(el));
  $.root().empty().append($main);

  // ---- Alerts: convert marker divs to <section class="alert alert-...">
  $("[data-wet^='alert-']").each((_, el) => {
    const type = $(el).attr("data-wet").replace("alert-", ""); // success/info/warning/danger
    const content = $(el).find("p").first();
    const $section = $("<section/>").addClass(`alert alert-${type}`);
    $section.append(content);
    $(el).replaceWith($section);
  });

  // ---- Tables: apply classes to the next <table> after a marker
  const tableClassMap = {
    "table-basic": "table",
    "table-striped": "table table-striped",
    "table-bordered": "table table-bordered",
    "table-hover": "table table-hover",
    "table-condensed": "table table-condensed",
  };

  // We walk markers in document order and find the next table
  $("[data-wet^='table-']").each((_, el) => {
    const kind = $(el).attr("data-wet"); // e.g. table-striped, table-responsive
    const $marker = $(el);

    // find next table after marker
    const $nextTable = $marker.nextAll("table").first();
    if ($nextTable.length) {
      if (kind === "table-responsive") {
        // wrap table
        const $wrap = $("<div/>").addClass("table-responsive");
        $nextTable.replaceWith($wrap.append($nextTable));
      } else if (tableClassMap[kind]) {
        $nextTable.attr("class", tableClassMap[kind]);
      }
    }
    $marker.remove(); // remove marker from output
  });

  // ---- Details blocks: Summary starts <details>, subsequent content lines add <p> until next non-content marker
  (function buildDetails() {
    const nodes = $main.contents().toArray();
    let i = 0;

    while (i < nodes.length) {
      const node = nodes[i];
      const $node = $(node);

      if ($node.is("div[data-wet='details-summary']")) {
        const summaryHtml = $node.html();
        const $details = $("<details/>");
        const $summary = $("<summary/>").html(summaryHtml);
        $details.append($summary);

        // consume following details-content nodes
        let j = i + 1;
        while (j < nodes.length && $(nodes[j]).is("div[data-wet='details-content']")) {
          const contentHtml = $(nodes[j]).html();
          $details.append($("<p/>").html(contentHtml));
          j++;
        }

        // replace nodes i..j-1 with details
        $node.replaceWith($details);
        for (let k = i + 1; k < j; k++) $(nodes[k]).remove();

        // refresh nodes snapshot after modifications
        const refreshed = $main.contents().toArray();
        nodes.splice(0, nodes.length, ...refreshed);
        i++; // continue
        continue;
      }

      i++;
    }

    // remove any stray markers
    $main.find("div[data-wet='details-content']").remove();
  })();

  // ---- Accordion: markers build <section class="wb-accordion"><details>...</details>...</section>
  (function buildAccordion() {
    const nodes = $main.contents().toArray();
    let i = 0;

    while (i < nodes.length) {
      const $n = $(nodes[i]);

      if ($n.is("div[data-wet='accordion-start']")) {
        const $section = $("<section/>").addClass("wb-accordion");
        let j = i + 1;

        let $currentItem = null;

        while (j < nodes.length && !$(nodes[j]).is("div[data-wet='accordion-end']")) {
          const $m = $(nodes[j]);

          if ($m.is("div[data-wet='accordion-heading']")) {
            // close previous item implicitly by starting a new one
            $currentItem = $("<details/>");
            $currentItem.append($("<summary/>").html($m.html()));
            $section.append($currentItem);
            $m.remove();
            j++;
            continue;
          }

          if ($m.is("div[data-wet='accordion-panel']")) {
            if (!$currentItem) {
              $currentItem = $("<details/>");
              $currentItem.append($("<summary/>").text("Details"));
              $section.append($currentItem);
            }
            $currentItem.append($("<p/>").html($m.html()));
            $m.remove();
            j++;
            continue;
          }

          // Any other node inside accordion block: treat as panel content if an item exists,
          // otherwise move it outside accordion (safer)
          if ($currentItem) {
            $currentItem.append($("<p/>").html($m.html ? $m.html() : $m.text()));
            $m.remove();
          } else {
            // move outside: stop accordion content
            break;
          }
          j++;
        }

        // remove start + end markers, insert accordion section
        $n.replaceWith($section);
        // remove the end marker if present
        if (j < nodes.length && $(nodes[j]).is("div[data-wet='accordion-end']")) {
          $(nodes[j]).remove();
        }

        // refresh nodes snapshot
        const refreshed = $main.contents().toArray();
        nodes.splice(0, nodes.length, ...refreshed);
        i++;
        continue;
      }

      i++;
    }

    // remove any stray accordion markers
    $main.find("div[data-wet^='accordion-']").remove();
  })();

  // ---- Pagination: markers build <nav><ul class="pagination">...</ul></nav>
  (function buildPagination() {
    const nodes = $main.contents().toArray();
    let i = 0;

    while (i < nodes.length) {
      const $n = $(nodes[i]);

      if ($n.is("div[data-wet='pagination-start']")) {
        const $nav = $("<nav/>").attr("aria-label", "Pagination");
        const $ul = $("<ul/>").addClass("pagination");
        $nav.append($ul);

        let j = i + 1;
        while (j < nodes.length && !$(nodes[j]).is("div[data-wet='pagination-end']")) {
          const $m = $(nodes[j]);

          if ($m.is("div[data-wet='pagination-active']")) {
            const label = $m.text().trim() || $m.html();
            const $li = $("<li/>").addClass("active");
            $li.append($("<a/>").attr("href", "#").attr("aria-current", "page").text(label));
            $ul.append($li);
            $m.remove();
            j++;
            continue;
          }

          if ($m.is("div[data-wet='pagination-disabled']")) {
            const label = $m.text().trim() || $m.html();
            const $li = $("<li/>").addClass("disabled");
            $li.append($("<span/>").text(label));
            $ul.append($li);
            $m.remove();
            j++;
            continue;
          }

          if ($m.is("div[data-wet='pagination-item']")) {
            const label = $m.text().trim() || $m.html();
            const $li = $("<li/>");
            $li.append($("<a/>").attr("href", "#").text(label));
            $ul.append($li);
            $m.remove();
            j++;
            continue;
          }

          // If some other node appears, stop pagination block (authoring error)
          break;
        }

        // replace start marker with nav, remove end marker if found
        $n.replaceWith($nav);
        if (j < nodes.length && $(nodes[j]).is("div[data-wet='pagination-end']")) {
          $(nodes[j]).remove();
        }

        // refresh nodes snapshot
        const refreshed = $main.contents().toArray();
        nodes.splice(0, nodes.length, ...refreshed);
        i++;
        continue;
      }

      i++;
    }

    // remove any stray pagination markers
    $main.find("div[data-wet^='pagination-']").remove();
  })();

    // ---- List markers: apply class to the next ul/ol after a marker
  $("[data-wet='list-inline'], [data-wet='list-unstyled']").each((_, el) => {
    const kind = $(el).attr("data-wet"); // list-inline or list-unstyled
    const $marker = $(el);

    const $nextList = $marker.nextAll("ul,ol").first();
    if ($nextList.length) {
      $nextList.addClass(kind);
    }

    $marker.remove();
  });

  return $.root().html();
}

async function convertDocxToHtml(inputPath) {
  const docxBuffer = fs.readFileSync(inputPath);

  const result = await mammoth.convertToHtml(
    { buffer: docxBuffer },
    {
      styleMap,
      // Keep mammoth’s HTML fairly clean
      includeDefaultStyleMap: true,
    }
  );

  // mammoth returns HTML fragment; wrap in a body for cheerio to parse reliably
  const raw = `<body>${result.value}</body>`;
  const finalHtml = transformPlaceholdersToGCWeb(raw);

  return { html: finalHtml, messages: result.messages };
}

(async function main() {
  const args = parseArgs(process.argv);
  if (!args.input) {
    console.error("Usage: node docx_to_gcweb_html_extended.js input.docx -o output.html");
    process.exit(2);
  }

  const inputPath = path.resolve(args.input);
  const { html, messages } = await convertDocxToHtml(inputPath);

  if (messages && messages.length) {
    // non-fatal warnings from mammoth
    console.error("mammoth messages:");
    for (const m of messages) console.error(" -", m.message || JSON.stringify(m));
  }

  if (args.output) {
    fs.writeFileSync(path.resolve(args.output), html, "utf8");
  } else {
    process.stdout.write(html);
  }
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
