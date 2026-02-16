// ============================================================================
// Parser_V3.gs
// Commodity Sampler Services — Custom Client Parsers
//
// Created: Feb 14, 2026
//   - Consolidated ALL client-specific parsers into one file
//   - Moved Serengeti + Atlantic out of EmailExtraction_v3.gs
//   - Added: Ally, Armenia, AMCOF, Coffee America, InterAmerican, ICC,
//            Osito, Rothfos, LDC, Sucafina/Tastify, Coffee Source
//   - Removed duplicate LDC definition
//
// DEPENDENCIES: EmailExtraction_v3.gs provides:
//   - detectWarehouse(), extractShipping(), extractFedExAccount()
//   - getOriginFromMark(), ICO_COUNTRIES
//   - extractSampleSize(), extractDescription()
// ============================================================================


// ============================================================================
// SERENGETI TRADING COMPANY
// ============================================================================

function extractSerengetiOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  // Strip Gmail forward markers
  allText = allText.replace(/>{2,}/g, ' ').replace(/\s+/g, ' ');

  // Shipping method
  var shipping = 'FedEx 2 Day';
  var shipMatch = allText.match(/Shipping Method:\s*([A-Za-z0-9\s-]+?)(?:Address:|$)/i);
  if (shipMatch) shipping = shipMatch[1].trim();

  // Receiver from Address block
  var receiver = 'Serengeti Trading Company';
  var addressBlock = allText.match(/Address:\s*([\s\S]*?)(?:Shipping Notes:|Quality:|$)/i);
  if (addressBlock) {
    var addrText = addressBlock[1].trim();
    var extracted = '';

    var attnMatch = addrText.match(/ATTN:\s*\S+\s+(.+?)\s+\d{3,}/i);
    if (attnMatch) {
      extracted = attnMatch[1].trim();
    } else {
      var beforeNumber = addrText.match(/^(.+?)\s+\d{3,}/);
      if (beforeNumber) extracted = beforeNumber[1].trim();
    }

    if (extracted) {
      var words = extracted.split(/\s+/);
      if (words.length >= 2) {
        var firstWord = words[0];
        var secondWord = words[1];
        var looksLikePerson = /^[A-Z][a-z]+$/.test(firstWord) &&
                              /^[A-Z][a-zA-Z]+$/.test(secondWord) &&
                              !/Trading|Company|Coffee|Commodities|Inc|LLC|Corp|Sourcing|Roasters|Services|Serengeti|Regal|Papa|Nicholas|Swift/i.test(firstWord) &&
                              !/Trading|Company|Coffee|Commodities|Inc|LLC|Corp|Sourcing|Roasters|Services|Swift/i.test(secondWord);

        if (looksLikePerson && words.length > 2) {
          receiver = words.slice(2).join(' ');
        } else {
          receiver = extracted;
        }
      } else {
        receiver = extracted;
      }

      if (receiver.split(/\s+/).length === 1) {
        var knownExpansions = {
          'Swift': 'Swift Coffee Sourcing',
          'Regal': 'Regal Commodities',
          'Papa': 'Papa Nicholas'
        };
        if (knownExpansions[receiver]) receiver = knownExpansions[receiver];
      }
    }
  }

  // Sender from signature
  var sender = '';
  var knownSenders = allText.match(/\b(Tyler\s+von\s+Roemer|Adriana\s+Imani)\b/i);
  if (knownSenders) {
    sender = knownSenders[1].replace(/\s+/g, ' ');
  }
  if (!sender) {
    var sigMatch = allText.match(/(?:Thank\s*[Yy]ou!?|[-–—]{2,})\s*([A-Z][a-z]+(?:\s+von)?\s+[A-Z][a-zA-Z]+)(?=\s+(?:Trading|Sales|Quality|Mobile|Office|The|Serengeti|$))/i);
    if (sigMatch) sender = sigMatch[1].trim();
  }

  // Split by "Quality:" to get each order block
  var blocks = allText.split(/Quality:/i);

  for (var i = 1; i < blocks.length; i++) {
    var block = blocks[i];

    if (/^\s*Reference:\s*Container:/i.test(block)) {
      var notesIdx = block.indexOf('Notes:');
      if (notesIdx > -1) block = block.substring(notesIdx + 6);
    }
    if (block.trim().length < 15) continue;

    var order = {
      client: 'Serengeti Trading Company',
      receiver: receiver,
      shipping: shipping,
      warehouse: 'Continental',
      sampleOrderNum: sender,
      container: '',
      mark: '',
      cargo: '',
      reference: '',
      description: '',
      bags: '',
      sampleSize: '4 lb',
      comments: '',
      origin: ''
    };

    // Container
    var containerMatch = block.match(/\b([A-Z]{4}\d{7})\b/);
    if (containerMatch) order.container = containerMatch[1];

    // ICO Mark
    var markMatch = block.match(/(\d{3}\/\d{3,4}\/\d+[A-Z]?)(?=C\d)/i);
    if (markMatch) {
      order.mark = markMatch[1];
    } else {
      markMatch = block.match(/\b(\d{3}\/\d{3,4}\/\d{3,6}[A-Z]?)\b/);
      if (markMatch) order.mark = markMatch[1];
    }

    // Cargo + Bags
    var cargoMatch = block.match(/(C\d{6})\s*(\d+)\s*(?:BGS?|bags?)/i);
    if (cargoMatch) {
      order.cargo = cargoMatch[1].toUpperCase();
      order.bags = cargoMatch[2];
    } else {
      var cargoOnly = block.match(/\b(C\d{6})\b/i);
      if (cargoOnly) order.cargo = cargoOnly[1].toUpperCase();
      var bagsMatch = block.match(/\b(\d+)\s*(?:BGS?|bags?)\b/i);
      if (bagsMatch) order.bags = bagsMatch[1];
    }

    // Reference
    var refMatch = block.match(/\b(P\d{5}[A-Z]?(?:-\d+)?)\s*(\([^)]+\))?/i);
    if (refMatch) {
      order.reference = refMatch[1].toUpperCase();
      if (refMatch[2]) order.reference += ' ' + refMatch[2];
    }
    if (!order.reference) {
      var ptRefMatch = block.match(/\b(PT\d{5}(?:-(?:S\d+|\d+))?)\b/i);
      if (ptRefMatch) order.reference = ptRefMatch[1].toUpperCase();
    }

    // PT as mark if no ICO mark
    if (!order.mark) {
      var ptMarkMatch = block.match(/\b(PT\d{5}(?:-\d+)?)\b/i);
      if (ptMarkMatch) order.mark = ptMarkMatch[1].toUpperCase();
    }

    // Sample size
    var weightMatch = block.match(/\b(\d+)\s*(?:lbs?)\b/i);
    if (weightMatch) order.sampleSize = weightMatch[1] + ' lb';

    // Description
    var description = '';
    var refPos = block.search(/P\d{5}|PT\d{5}/i);
    if (refPos > 0) {
      var beforeRef = block.substring(0, refPos).trim();
      beforeRef = beforeRef.replace(/.*Notes:\s*/i, '');
      beforeRef = beforeRef.replace(/.*Weight:\s*/i, '');
      beforeRef = beforeRef.replace(/.*Location:\s*/i, '');
      beforeRef = beforeRef.replace(/.*Quantity:\s*/i, '');
      beforeRef = beforeRef.trim();
      if (beforeRef.length > 3 && /^[A-Za-z]/.test(beforeRef)) description = beforeRef;
    }
    if (description) order.description = description;

    // Origin from mark — uses shared ICO_COUNTRIES lookup
    if (order.mark && /^\d{3}\//.test(order.mark)) {
      order.origin = getOriginFromMark(order.mark);
    }

    // Only use origin as description if nothing else found
    if (!order.description && order.origin) order.description = order.origin;

    // Comments
    var wetMatch = block.match(/(\d+\s*Wet\s*Bags?)/i);
    if (wetMatch) order.comments = wetMatch[1];

    if (order.cargo || order.container) orders.push(order);
  }

  return orders;
}


// ============================================================================
// ATLANTIC (USA), LLC
// ============================================================================

function extractAtlanticOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  // Reference: "LETTER OF ENTRY AP30277G"
  var reference = '';
  var refMatch = allText.match(/(?:LETTER\s+OF\s+ENTRY|Our\s+Reference)\s*[:\s]*([A-Z]{2}\d{4,8}[A-Z]?)/i);
  if (refMatch) reference = refMatch[1].toUpperCase();

  var warehouse = detectWarehouse(allText) || 'RPM';

  // Sample size: "draw one 2 lb sample per cntr"
  var sampleSize = '2 lb';
  var ssMatch = allText.match(/draw\s+(?:one\s+)?(\d+)\s*(lb|lbs|g)\s+sample/i);
  if (ssMatch) {
    var u = (ssMatch[2] || 'lb').toLowerCase();
    sampleSize = u.startsWith('g') ? ssMatch[1] + 'g' : ssMatch[1] + ' lb';
  }

  // Shipping
  var shipping = 'UPS Ground';
  if (/UPS\s+GROUND/i.test(allText)) shipping = 'UPS Ground';
  else if (/UPS\s+2nd/i.test(allText)) shipping = 'UPS 2nd Day';
  else if (/fedex\s+standard\s+overnight/i.test(allText)) shipping = 'FedEx Standard Overnight';
  else shipping = extractShipping(allText, emailBody) || 'UPS Ground';

  // Receiver
  var receiver = 'Atlantic (USA), LLC';
  var recvMatch = allText.match(/send\s+to\s+([A-Za-z][A-Za-z\s(,).]+?)(?:\s+via\b)/i);
  if (recvMatch) receiver = recvMatch[1].trim();

  // Description from Quality field
  var description = '';
  var qualMatch = allText.match(/Quality\s*[:\s]+([^\n]+)/i);
  if (qualMatch) description = qualMatch[1].trim();

  // Parse table lines with containers
  var lines = allText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    var contMatch = line.match(/\b([A-Z]{4}\d{7})\b/);
    if (!contMatch) continue;
    var container = contMatch[1];

    // Skip headers and instruction lines
    if (/\bContainer\b/i.test(line) && /\bSeals?\b/i.test(line) && /\bMarks?\b/i.test(line)) continue;
    if (/\bTotal\s+Quantity\b/i.test(line)) continue;
    if (/delivered\s+to|pick\s+up|warehouse|trucker|rpm_inbound/i.test(line)) continue;

    // Mark: dash-separated, 3 or 4 parts → normalize to slash format
    var mark = '';
    var markMatch = line.match(/\b(\d{1,3})\s*[-]\s*(\d{2,5})\s*[-]\s*(\d{3,6}[A-Z]?)(?:\s*[-]\s*\d{1,4})?\b/);
    if (markMatch) {
      var cc = markMatch[1];
      while (cc.length < 3) cc = '0' + cc;
      mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
    }

    // Bags: number after mark
    var bags = '';
    if (markMatch) {
      var afterMark = line.substring(line.indexOf(markMatch[0]) + markMatch[0].length);
      var bagsMatch = afterMark.match(/\b(\d{1,4})\b/);
      if (bagsMatch && parseInt(bagsMatch[1]) <= 2000) bags = bagsMatch[1];
    }

    // Lot Ref per row
    var lotRef = reference;
    if (markMatch) {
      var afterMark2 = line.substring(line.indexOf(markMatch[0]) + markMatch[0].length);
      var lotMatch = afterMark2.match(/\b([A-Z]{2}\d{4,8}[A-Z]?)\b/);
      if (lotMatch) lotRef = lotMatch[1];
    }

    var origin = getOriginFromMark(mark);

    orders.push({
      client: 'Atlantic (USA), LLC',
      sampleOrderNum: '',
      container: container,
      mark: mark,
      cargo: '',  // Atlantic LOEs have NO cargo numbers
      reference: lotRef || reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: bags,
      origin: origin,
      description: description || origin || '',
      comments: ''
    });
  }

  return orders;
}


// ============================================================================
// ALLY COFFEE TRADING S.A.
// ============================================================================

function extractAllyOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  // Order number
  var orderNum = '';
  var orderMatch = allText.match(/Sampling Order\s*#[:\s]*(\d+)/i);
  if (orderMatch) orderNum = orderMatch[1];

  // Reference: P006559
  var reference = '';
  var refMatch = allText.match(/Our Reference[:\s]*(P\d{5,9}[A-Z]?)/i);
  if (refMatch) reference = refMatch[1];

  var warehouse = detectWarehouse(allText) || 'RPM';

  // Sample Size
  var sampleSize = '2 lb';
  var sizeMatch = allText.match(/Sample Size[:\s]*(\d+)\s*(LB|lb|lbs|g)/i);
  if (sizeMatch) {
    var unit = (sizeMatch[2] || 'lb').toLowerCase();
    sampleSize = sizeMatch[1] + (unit.startsWith('g') ? 'g' : ' lb');
  }

  // Receiver
  var receiver = '';
  var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
  if (recvMatch) {
    receiver = recvMatch[1].trim();
    receiver = receiver.replace(/\d{3,}.*$/, '').trim();
    receiver = receiver.replace(/,\s*$/, '').trim();
  }
  if (!receiver) receiver = 'Ally Coffee Trading S.A.';

  // Shipping
  var shipping = extractShipping(allText, emailBody) || 'FedEx Standard Overnight';

  // Parse table rows by container
  var containerPattern = /\b([A-Z]{4}\d{7})\b/g;
  var match;
  var rowsProcessed = {};

  while ((match = containerPattern.exec(allText)) !== null) {
    var container = match[1];
    if (rowsProcessed[container]) continue;
    rowsProcessed[container] = true;

    var start = Math.max(0, match.index - 200);
    var end = Math.min(allText.length, match.index + container.length + 100);
    var context = allText.substring(start, end);

    var lines = context.split('\n');
    var containerLine = '';
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].indexOf(container) >= 0) {
        containerLine = lines[i];
        break;
      }
    }

    // Bag count
    var bags = '';
    var bagsMatch = containerLine.match(/\b(\d{2,4})\s+\d{4,6}\s+BAGS\b/i);
    if (bagsMatch) {
      bags = bagsMatch[1];
    } else {
      var bagsFallback = containerLine.match(/\b(\d{2,4})\b/);
      if (bagsFallback && parseInt(bagsFallback[1]) < 1000) bags = bagsFallback[1];
    }

    // Skip "Total" rows
    if (/\bTotal\b/i.test(containerLine)) continue;

    // ICO Mark
    var mark = '';
    var markMatch = containerLine.match(/\b(\d{3})\s*\/\s*(\d{3,5})\s*\/\s*(\d{3,6}[A-Z]?)\b/);
    if (markMatch) mark = markMatch[1] + '/' + markMatch[2] + '/' + markMatch[3];

    // Cargo
    var cargo = '';
    var afterContainer = containerLine.substring(containerLine.indexOf(container) + container.length);
    var cargoMatch = afterContainer.match(/\b(\d{6})\b/);
    if (cargoMatch) cargo = 'C' + cargoMatch[1];
    if (!cargo) {
      var cCargoMatch = containerLine.match(/\b(C\d{6})\b/i);
      if (cCargoMatch) cargo = cCargoMatch[1].toUpperCase();
    }

    // Item# reference
    var itemRef = reference;
    var itemMatch = afterContainer.match(/\b(P\d{5,9}(?:#\d+)?)\b/i);
    if (itemMatch) itemRef = itemMatch[1].replace(/#\d+$/, '');

    var origin = mark ? getOriginFromMark(mark) : '';

    orders.push({
      client: 'Ally Coffee Trading S.A.',
      sampleOrderNum: orderNum,
      container: container,
      mark: mark,
      cargo: cargo,
      reference: itemRef || reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: bags,
      origin: origin,
      description: origin || '',
      comments: ''
    });
  }

  // FedEx account
  var fedex = extractFedExAccount(allText);
  if (fedex) {
    for (var j = 0; j < orders.length; j++) orders[j].fedexAccount = fedex;
  }

  return orders;
}


// ============================================================================
// ARMENIA COFFEE CORP
// ============================================================================

function extractArmeniaOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var orderNum = '';
  var orderMatch = allText.match(/Sampling Order\s*#[:\s]*(\d+)/i);
  if (orderMatch) orderNum = orderMatch[1];

  // Reference: S239J4-1
  var reference = '';
  var refMatch = allText.match(/Our Reference[:\s]*(S\d+[A-Z0-9-]+)/i);
  if (refMatch) reference = refMatch[1];
  if (!reference) {
    var custRefMatch = allText.match(/Customers?\s+Reference\s*#?[:\s]*([A-Z0-9-]+)/i);
    if (custRefMatch) reference = custRefMatch[1];
  }

  var warehouse = detectWarehouse(allText) || 'Continental';

  // Sample size — check Remarks first
  var sampleSize = '2 lb';
  var remarksSizeMatch = allText.match(/(?:send|draw|pull)\s+(?:a\s+)?(\d+)\s*(lb|lbs|g|gram)\s+sample/i);
  if (remarksSizeMatch) {
    var unit = (remarksSizeMatch[2] || 'lb').toLowerCase();
    sampleSize = remarksSizeMatch[1] + (unit.startsWith('g') ? 'g' : ' lb');
  } else {
    var sizeMatch = allText.match(/Sample Size[:\s]*(\d+)\s*(LB|lb|lbs|g)/i);
    if (sizeMatch) {
      var unit2 = (sizeMatch[2] || 'lb').toLowerCase();
      sampleSize = sizeMatch[1] + (unit2.startsWith('g') ? 'g' : ' lb');
    }
  }

  // Receiver
  var receiver = '';
  var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
  if (recvMatch) {
    receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  }
  if (!receiver) receiver = 'Armenia Coffee Corp';

  // Sender
  var sender = '';
  var senderMatch = allText.match(/\n([A-Z][a-z]+\s+[A-Z][a-z]+)\s*\n\s*Armenia\s+Coffee/i);
  if (senderMatch) sender = senderMatch[1];
  if (!sender) sender = 'Armenia Coffee Corp';

  // Shipping — check Remarks
  var shipping = '';
  var shippingFromRemarks = allText.match(/(?:sample|send)\s+(?:via\s+)?(?:UPS|FedEx)[^\n]*/i);
  if (shippingFromRemarks) {
    var shText = shippingFromRemarks[0];
    if (/UPS\s+Ground/i.test(shText)) shipping = 'UPS Ground';
    else if (/UPS\s+2nd?\s+Day/i.test(shText)) shipping = 'UPS 2nd Day';
    else if (/UPS/i.test(shText)) shipping = 'UPS';
    else if (/FedEx\s+Ground/i.test(shText)) shipping = 'FedEx Ground';
    else if (/FedEx\s+2[\s-]?Day/i.test(shText)) shipping = 'FedEx 2 Day';
    else if (/FedEx/i.test(shText)) shipping = 'FedEx';
  }
  if (!shipping) shipping = extractShipping(allText, emailBody) || 'UPS Ground';

  // Bags from Total Quantity
  var totalBags = '';
  var totalMatch = allText.match(/Total\s+Quantity[:\s]*(\d+)\s*BAGS/i);
  if (totalMatch) totalBags = totalMatch[1];

  // Description from Total Quantity line
  var description = '';
  var descMatch = allText.match(/Total\s+Quantity[:\s]*\d+\s*BAGS\s+(.+)/i);
  if (descMatch) {
    description = descMatch[1].trim().replace(/\s+Crop\s+\d{2}\/\d{2}\s*$/, '').trim();
  }

  // Parse table rows by container
  var containerPattern = /\b([A-Z]{4}\d{7})\b/g;
  var match;
  var rowsProcessed = {};

  while ((match = containerPattern.exec(allText)) !== null) {
    var container = match[1];
    if (rowsProcessed[container]) continue;
    rowsProcessed[container] = true;

    var lineStart = allText.lastIndexOf('\n', match.index) + 1;
    var lineEnd = allText.indexOf('\n', match.index + container.length);
    var line = allText.substring(lineStart, lineEnd > 0 ? lineEnd : allText.length);

    if (/\bContainer\s+No\b/i.test(line) && !/[A-Z]{4}\d{7}/.test(line.replace(/Container\s+No/i, ''))) continue;

    // Cargo
    var cargo = '';
    var cargoMatch = line.match(/\b(C\d{5,7}[A-Z]?)\b/i);
    if (cargoMatch) cargo = cargoMatch[1].toUpperCase();

    // ICO Mark (handles short country codes)
    var mark = '';
    var markMatch = line.match(/\b(\d{1,3})\s*\/\s*(\d{2,5})\s*\/\s*(\d{3,6}[A-Z]?)\b/);
    if (markMatch) {
      var countryCode = markMatch[1];
      while (countryCode.length < 3) countryCode = '0' + countryCode;
      mark = countryCode + '/' + markMatch[2] + '/' + markMatch[3];
    }

    // Bags from row
    var rowBags = '';
    var afterMark = mark ? line.substring(line.indexOf(markMatch[0]) + markMatch[0].length) : '';
    var unitsMatch = afterMark.match(/\b(\d+)(?:\.\d+)?\b/);
    if (unitsMatch) rowBags = unitsMatch[1];
    if (!rowBags) rowBags = totalBags;

    // Row description
    var rowDesc = description;
    var descFromRow = afterMark.match(/\d+(?:\.\d+)?\s+(.+)/);
    if (descFromRow && descFromRow[1].trim().length > 5) rowDesc = descFromRow[1].trim();

    var origin = mark ? getOriginFromMark(mark) : '';

    var order = {
      client: 'Armenia Coffee Corp',
      sampleOrderNum: orderNum,
      container: container,
      mark: mark,
      cargo: cargo,
      reference: reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: rowBags,
      origin: origin,
      description: rowDesc || origin || '',
      comments: ''
    };

    var conditionMatch = allText.match(/Recondition[^\n]*/i);
    if (conditionMatch) order.comments = conditionMatch[0].trim();

    orders.push(order);
  }

  return orders;
}


// ============================================================================
// AMERICAN COFFEE CORPORATION (AMCOF)
// ============================================================================

function extractAmcofOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var orderNum = '';
  var orderMatch = allText.match(/Sampling Order\s*#[:\s]*(\d+)/i);
  if (orderMatch) orderNum = orderMatch[1];

  // Reference: P017792
  var reference = '';
  var refMatch = allText.match(/Our Reference[:\s]*(P\d{5,9}[A-Z]?)/i);
  if (refMatch) reference = refMatch[1];
  if (!reference) {
    var purchMatch = allText.match(/\b(P?\d{5,7})\s+\d{3,4}\/\d{3,4}/);
    if (purchMatch) {
      var purch = purchMatch[1];
      if (purch.charAt(0) !== 'P') purch = 'P' + purch;
      reference = purch;
    }
  }

  var warehouse = detectWarehouse(allText) || 'Continental';

  var sampleSize = '2 lb';
  var remarksSizeMatch = allText.match(/(?:send|draw|pull)\s+(?:a\s+)?(\d+)\s*(lb|lbs|g|gram)\s+sample/i);
  if (remarksSizeMatch) {
    var unit = (remarksSizeMatch[2] || 'lb').toLowerCase();
    sampleSize = remarksSizeMatch[1] + (unit.startsWith('g') ? 'g' : ' lb');
  } else {
    var sizeMatch = allText.match(/Sample Size[:\s]*(\d+)\s*(LB|lb|lbs|g)/i);
    if (sizeMatch) sampleSize = sizeMatch[1] + ' lb';
  }

  var receiver = '';
  var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
  if (recvMatch) {
    receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  }
  if (!receiver) receiver = 'American Coffee Corp.';

  var shipping = extractShipping(allText, emailBody) || 'FedEx';

  var totalBags = '';
  var totalMatch = allText.match(/Total\s+Quantity[:\s]*(\d+)\s*Bags/i);
  if (totalMatch) totalBags = totalMatch[1];

  var description = '';
  var descMatch = allText.match(/Total\s+Quantity[:\s]*\d+\s*Bags\s+(.+)/i);
  if (descMatch) description = descMatch[1].trim();

  var containerPattern = /\b([A-Z]{4}\d{7})\b/g;
  var match;
  var rowsProcessed = {};

  while ((match = containerPattern.exec(allText)) !== null) {
    var container = match[1];
    if (rowsProcessed[container]) continue;
    rowsProcessed[container] = true;

    var lineStart = allText.lastIndexOf('\n', match.index) + 1;
    var lineEnd = allText.indexOf('\n', match.index + container.length);
    var line = allText.substring(lineStart, lineEnd > 0 ? lineEnd : allText.length);

    if (/\bContainer\s+No\b/i.test(line) && /\bPurch\b/i.test(line)) continue;

    var mark = '';
    var markMatch = line.match(/\b(\d{1,4})\s*\/\s*(\d{2,5})\s*\/\s*(\d{3,6}[A-Z]?)\b/);
    if (!markMatch) {
      var beforeContainer = allText.substring(Math.max(0, match.index - 100), match.index);
      markMatch = beforeContainer.match(/\b(\d{1,4})\s*\/\s*(\d{2,5})\s*\/\s*(\d{3,6}[A-Z]?)\b/);
    }
    if (markMatch) {
      var countryCode = markMatch[1];
      if (countryCode.length === 4 && countryCode.charAt(0) === '0') countryCode = countryCode.substring(1);
      while (countryCode.length < 3) countryCode = '0' + countryCode;
      mark = countryCode + '/' + markMatch[2] + '/' + markMatch[3];
    }

    var cargo = '';
    var cargoMatch = line.match(/\b(C\d{5,7}[A-Z]?)\b/i);
    if (cargoMatch) cargo = cargoMatch[1].toUpperCase();

    var rowBags = '';
    var afterContainer = line.substring(line.indexOf(container) + container.length);
    var unitsMatch = afterContainer.match(/\b(\d{2,4})\b/);
    if (unitsMatch && parseInt(unitsMatch[1]) < 2000) rowBags = unitsMatch[1];
    if (!rowBags) rowBags = totalBags;

    var origin = mark ? getOriginFromMark(mark) : '';

    orders.push({
      client: 'American Coffee Corporation',
      sampleOrderNum: orderNum,
      container: container,
      mark: mark,
      cargo: cargo,
      reference: reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: rowBags,
      origin: origin,
      description: description || origin || '',
      comments: ''
    });
  }

  return orders;
}


// ============================================================================
// COFFEE AMERICA USA CORP
// ============================================================================

function extractCoffeeAmericaOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var orderNum = '';
  var orderMatch = allText.match(/(?:Warehouse\s+)?Sampling\s+Order\s*#[:\s]*0*(\d+)/i);
  if (orderMatch) orderNum = orderMatch[1];
  if (!orderNum) {
    var altMatch = allText.match(/SAMPLE\s+ORDER\s*#\s*(\d+)/i);
    if (altMatch) orderNum = altMatch[1];
  }

  var headerRef = '';
  var refMatch = allText.match(/Our Ref\s*#?[:\s]*(P\d{5}[\d.]+)/i);
  if (refMatch) headerRef = refMatch[1];

  var warehouse = detectWarehouse(allText) || 'RPM';

  var sampleSize = '2 lb';
  var sizeMatch = allText.match(/(?:send|draw|pull)\s+(?:a\s+)?(\d+)\s*(lb|lbs|g|gram)\s+sample/i);
  if (sizeMatch) {
    sampleSize = sizeMatch[1] + ((sizeMatch[2] || 'lb').toLowerCase().startsWith('g') ? 'g' : ' lb');
  }

  // Receiver: "FOR:" field
  var receiver = '';
  var forMatch = allText.match(/\bFOR[:\s]+([A-Za-z][A-Za-z\s&'.,-]+?)(?=\n|\d{3,})/i);
  if (forMatch) receiver = forMatch[1].trim().replace(/,\s*$/, '').trim();
  if (!receiver) {
    var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
    if (recvMatch) receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  }
  if (!receiver) receiver = 'Coffee America (USA) Corporation';

  // Shipping
  var shipping = 'FedEx Priority';
  if (/priority\s+Fed/i.test(allText)) shipping = 'FedEx Priority';
  else if (/standard\s+overnight/i.test(allText)) shipping = 'FedEx Standard Overnight';
  else if (/fedex\s+ground/i.test(allText)) shipping = 'FedEx Ground';
  else if (/fedex\s+2[\s-]?day/i.test(allText)) shipping = 'FedEx 2 Day';
  else shipping = extractShipping(allText, emailBody) || 'FedEx Priority';

  // FedEx account
  var fedexAccount = '';
  var fedexMatch = allText.match(/account\s+number\s+(\d{4}[-]?\d{4}[-]?\d{1,3})/i);
  if (fedexMatch) fedexAccount = fedexMatch[1];
  if (!fedexAccount) fedexAccount = extractFedExAccount(allText) || '6701-4695-2';

  var description = '';
  var descMatch = allText.match(/Total\s+Quantity[:\s]*\d+\s+(?:\d+\s+KG\s+)?Bags\s+(.+)/i);
  if (descMatch) description = descMatch[1].trim();

  // Parse LINE BY LINE (same container can appear multiple times)
  var lines = allText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    var containerMatch = line.match(/\b([A-Z]{4}\d{7})\b/);
    if (!containerMatch) continue;
    var container = containerMatch[1];

    if (/\bContainer\s+No\b/i.test(line) && /\bICO\s+Marks?\b/i.test(line)) continue;
    if (/\bTotal\b/i.test(line) && !/REC/i.test(line) && !/P\d{5}/i.test(line)) continue;

    // Line reference
    var lineRef = '';
    var lineRefMatch = line.match(/\b(P\d{5}[\d.]+)\b/i);
    if (lineRefMatch) lineRef = lineRefMatch[1];
    if (!lineRef) lineRef = headerRef;

    // ICO Mark
    var mark = '';
    var markMatch = line.match(/\b(\d{1,4})\s*\/\s*(\d{2,5})\s*\/\s*(\d{3,6}[A-Z]?)\b/);
    if (markMatch) {
      var cc = markMatch[1];
      if (cc.length === 4 && cc.charAt(0) === '0') cc = cc.substring(1);
      while (cc.length < 3) cc = '0' + cc;
      mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
    }

    // Warrant (stored as cargo)
    var cargo = '';
    if (markMatch) {
      var afterMark = line.substring(line.indexOf(markMatch[0]) + markMatch[0].length, line.indexOf(container));
      var warrantInBetween = afterMark.match(/\b(\d{5,7}[A-Z]?)\b/);
      if (warrantInBetween) cargo = warrantInBetween[1];
    }

    // Bags + REC comments
    var bags = '';
    var comments = '';
    var afterContainer = line.substring(line.indexOf(container) + container.length);
    var bagsMatch = afterContainer.match(/\b(\d{1,4})\s*(REC)?/i);
    if (bagsMatch) {
      bags = bagsMatch[1];
      if (bagsMatch[2]) comments = bags + ' REC (Reconditioned)';
    }

    var origin = getOriginFromMark(mark);

    orders.push({
      client: 'Coffee America USA Corp',
      sampleOrderNum: orderNum,
      container: container,
      mark: mark,
      cargo: cargo,
      reference: lineRef,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: fedexAccount,
      bags: bags,
      origin: origin,
      description: description || origin || '',
      comments: comments
    });
  }

  return orders;
}


// ============================================================================
// INTERAMERICAN COFFEE (NKG)
// ============================================================================

function extractInterAmericanOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  // Reference: "Ctrt # / Shipment #: P200178 / 1"
  var reference = '';
  var refMatch = allText.match(/Ctrt\s*#?\s*\/\s*Shipment\s*#?[:\s]*(P\d{5,9})/i);
  if (refMatch) reference = refMatch[1];
  if (!reference) {
    var pMatch = allText.match(/\b(P\d{5,9})\b/);
    if (pMatch) reference = pMatch[1];
  }

  var warehouse = detectWarehouse(allText) || 'Continental';

  // Sample size: "Sample Qty: 3Lbs" or email "send 3lbs"
  var sampleSize = '2 lb';
  var sqMatch = allText.match(/Sample\s+Qty[:\s]*(\d+)\s*(Lbs?|lb|g)/i);
  if (sqMatch) {
    var unit = (sqMatch[2] || 'lb').toLowerCase();
    sampleSize = sqMatch[1] + (unit.startsWith('g') ? 'g' : ' lb');
  } else {
    var emailMatch = allText.match(/send\s+(\d+)\s*(lbs?|lb|g)\b/i);
    if (emailMatch) sampleSize = emailMatch[1] + ' lb';
  }

  // Container
  var container = '';
  var contMatch = allText.match(/Container\s*#?[:\s]*([A-Z]{4}\d{7})/i);
  if (contMatch) container = contMatch[1];
  if (!container) {
    var anyContMatch = allText.match(/\b([A-Z]{4}\d{7})\b/);
    if (anyContMatch) container = anyContMatch[1];
  }

  // Shipment Marks (dash-separated)
  var mark = '';
  var markMatch = allText.match(/Shipment\s+Marks?[:\s]*(\d{1,3})\s*[-]\s*(\d{2,5})\s*[-]\s*(\d{3,6}[A-Z]?)/i);
  if (markMatch) {
    var cc = markMatch[1];
    while (cc.length < 3) cc = '0' + cc;
    mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
  }
  if (!mark) {
    var slashMark = allText.match(/(?:Marks?|ICO)[:\s]*(\d{1,3})\s*\/\s*(\d{2,5})\s*\/\s*(\d{3,6}[A-Z]?)/i);
    if (slashMark) {
      var cc2 = slashMark[1];
      while (cc2.length < 3) cc2 = '0' + cc2;
      mark = cc2 + '/' + slashMark[2] + '/' + slashMark[3];
    }
  }

  // Cargo
  var cargo = '';
  var cargoMatch = allText.match(/Cargo\s+(?:Ref\s+)?#?[:\s]*(C\d{5,7}[A-Z]?)/i);
  if (cargoMatch) cargo = cargoMatch[1].toUpperCase();
  if (!cargo) {
    var cMatch = allText.match(/\b(C\d{5,7})\b/i);
    if (cMatch) cargo = cMatch[1].toUpperCase();
  }

  // Bags — skip weight-based Available Qty
  var bags = '';
  var bagsMatch = allText.match(/Available\s+Qty[:\s]*(\d+)\s+(Kg\s+)?Bags/i);
  if (bagsMatch && !bagsMatch[2]) bags = bagsMatch[1];
  if (!bags) {
    var ohMatch = allText.match(/\b(\d+)\s+BAGS\b/i);
    if (ohMatch && !/Kg/i.test(allText.substring(Math.max(0, ohMatch.index - 10), ohMatch.index))) {
      bags = ohMatch[1];
    }
  }

  // Description
  var description = '';
  var originMatch = allText.match(/Origin\s+Grade[^\n]*\n\s*([^\n]+)/i);
  if (originMatch) description = originMatch[1].trim().replace(/\s+R\d{8,}.*$/, '').trim();

  // Receiver
  var receiver = '';
  var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
  if (recvMatch) receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  if (!receiver) receiver = 'InterAmerican Coffee';

  var shipping = extractShipping(allText, emailBody) || 'FedEx 2 Day';
  var origin = getOriginFromMark(mark);

  if (container || mark || cargo) {
    orders.push({
      client: 'InterAmerican Coffee',
      sampleOrderNum: '',
      container: container,
      mark: mark,
      cargo: cargo,
      reference: reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: bags,
      origin: origin,
      description: description || origin || '',
      comments: ''
    });
  }

  return orders;
}


// ============================================================================
// INTERNATIONAL COFFEE CORPORATION (ICC)
// ============================================================================

function extractIccOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var orderNum = '';
  var orderMatch = allText.match(/Sample\s+Order\s*#?[:\s]*(\d+)/i);
  if (orderMatch) orderNum = orderMatch[1];

  var warehouse = detectWarehouse(allText) || 'Continental';

  // Sample size: "Send *1* LB Sample(s)"
  var sampleSize = '1 lb';
  var sizeMatch = allText.match(/Send\s+\*?(\d+)\*?\s*(LB|lb|lbs|g)\s+Sample/i);
  if (sizeMatch) sampleSize = sizeMatch[1] + ' lb';

  // Receiver
  var receiver = '';
  var recvMatch = allText.match(/(?:Please\s+)?Send\s+Samples?\s+[Tt]o[:\s]+([^\n]+)/i);
  if (recvMatch) receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  if (!receiver || receiver.toLowerCase() === 'icc') {
    var belowMatch = allText.match(/send\s+sample\s+to\s+the\s+below\s+address[:\s]*\n\s*([A-Za-z][^\n]+)/i);
    if (belowMatch) receiver = belowMatch[1].trim();
  }
  if (!receiver) receiver = 'ICC';

  // Shipping
  var shipping = '';
  if (/2nd\s*Day/i.test(allText)) shipping = 'FedEx 2 Day';
  else if (/standard\s+overnight/i.test(allText)) shipping = 'FedEx Standard Overnight';
  else if (/priority/i.test(allText)) shipping = 'FedEx Priority';
  else if (/ground/i.test(allText)) shipping = 'FedEx Ground';
  else shipping = extractShipping(allText, emailBody) || 'FedEx Standard Overnight';

  // Parse PDF table rows (standard + ICC-style dashed containers)
  var lines = allText.split('\n');
  var processedLines = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    var contMatch = line.match(/\b([A-Z]{4}[-]?\d{6,7}[-]?\d?)\b/);
    if (!contMatch) continue;

    var rawContainer = contMatch[1];
    if (/\bContainer\b/i.test(line) && /\bDescription\b|Marks\b|Cargo\b/i.test(line)) continue;
    if (/Phone|Fax|Toll|Email|FLO|FTUSA|FT USA/i.test(line)) continue;

    var container = rawContainer.replace(/-/g, '');

    var lineKey = line.trim();
    if (processedLines[lineKey]) continue;
    processedLines[lineKey] = true;

    // ICO Mark
    var mark = '';
    var markMatch = line.match(/\b(\d{1,3})\s*[-\/]\s*(\d{2,5})\s*[-\/]\s*(\d{3,6}[A-Z]?)\b/);
    if (markMatch) {
      if (!/^20[0-3]\d$/.test(markMatch[1])) {
        var cc = markMatch[1];
        while (cc.length < 3) cc = '0' + cc;
        mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
      }
    }

    var cargo = '';
    var cargoMatch = line.match(/\b[Cc](\d{5,7}[A-Z]?)\b/);
    if (cargoMatch) cargo = 'C' + cargoMatch[1];

    var bags = '';
    var bagsMatch = line.match(/\b(\d{2,4})\s+BAGS\b/i);
    if (bagsMatch) bags = bagsMatch[1];

    var iccRef = '';
    var iccRefMatch = line.match(/\b(\d{4,6}[A-Z]?)\s*$/);
    if (iccRefMatch) iccRef = iccRefMatch[1];

    var lineDesc = '';
    var descMatch = line.match(/^([A-Z][A-Z\s]+?(?:ARABICA|COFFEE|ORGANIC|WASHED|NATURAL|EP|TP|SHB|SHG))\s+\d/i);
    if (descMatch) lineDesc = descMatch[1].trim();

    var origin = getOriginFromMark(mark);

    orders.push({
      client: 'International Coffee Corporation',
      sampleOrderNum: orderNum,
      container: container,
      mark: mark,
      cargo: cargo,
      reference: iccRef || '',
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: bags,
      origin: origin,
      description: lineDesc || origin || '',
      comments: ''
    });
  }

  // Parse inline email orders (Jude format)
  if (orders.length === 0 && emailBody) {
    var emailLines = emailBody.split('\n');
    for (var j = 0; j < emailLines.length; j++) {
      var eLine = emailLines[j].trim();
      var eContMatch = eLine.match(/\b([A-Z]{4}[-]?\d{6,7}[-]?\d?)\b/);
      if (!eContMatch) continue;
      if (/Phone|Fax|Toll|Email|FLO|FTUSA/i.test(eLine)) continue;

      var eContainer = eContMatch[1].replace(/-/g, '');

      var eCargo = '';
      var eCargoM = eLine.match(/\b[Cc](\d{5,7})\b/);
      if (eCargoM) eCargo = 'C' + eCargoM[1];

      var beforeCont = eLine.substring(0, eLine.indexOf(eContMatch[0])).trim();
      var markParts = beforeCont.match(/\b(\d{1,3})\s+(\d{2,5})\s+(\d{3,6})\b/);
      var eMark = '';
      if (markParts) {
        var ecc = markParts[1];
        while (ecc.length < 3) ecc = '0' + ecc;
        eMark = ecc + '/' + markParts[2] + '/' + markParts[3];
      }

      var eRef = '';
      if (j > 0) {
        var prevLine = emailLines[j - 1].trim();
        if (/^\d{4,6}[A-Z]?$/.test(prevLine)) eRef = prevLine;
      }

      var eShipping = shipping;
      if (/standard\s+overnight/i.test(emailBody)) eShipping = 'FedEx Standard Overnight';
      else if (/2nd\s*day/i.test(emailBody)) eShipping = 'FedEx 2 Day';

      var eOrigin = getOriginFromMark(eMark);

      orders.push({
        client: 'International Coffee Corporation',
        sampleOrderNum: '',
        container: eContainer,
        mark: eMark,
        cargo: eCargo,
        reference: eRef,
        warehouse: warehouse,
        receiver: receiver || 'ICC',
        sampleSize: sampleSize,
        shipping: eShipping,
        fedexAccount: '',
        bags: '',
        origin: eOrigin,
        description: eOrigin || '',
        comments: 'From inline email'
      });
    }
  }

  return orders;
}


// ============================================================================
// OSITO COFFEE
// ============================================================================

function extractOsitoOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var warehouse = detectWarehouse(allText) || 'Continental';

  var receiver = '';
  var recvMatch = allText.match(/send\s+(?:them\s+)?(?:via\s+\S+\s+)?(?:\([^)]*\)\s+)?to[:\s]*\n?\s*([^\n]+)/i);
  if (recvMatch) receiver = recvMatch[1].trim();
  if (!receiver) {
    var ositoRecv = allText.match(/Osito\s+Coffee\s+([A-Z][a-z]+\s+[A-Z][a-z]+)/i);
    if (ositoRecv) receiver = 'Osito Coffee - ' + ositoRecv[1];
  }
  if (!receiver) receiver = 'Osito Coffee';

  var shipping = 'FedEx Ground';
  if (/fedex\s+ground/i.test(allText)) shipping = 'FedEx Ground';
  else if (/standard\s+overnight/i.test(allText)) shipping = 'FedEx Standard Overnight';
  else shipping = extractShipping(allText, emailBody) || 'FedEx Ground';

  var fedexAccount = '';
  var fxMatch = allText.match(/account\s*[-:]?\s*(\d{7,12})/i);
  if (fxMatch) fedexAccount = fxMatch[1];

  // Parse cargo-based items
  var cargoPattern = /\b(C\d{5,7})\s*/gi;
  var cargoMatches = [];
  var match;
  while ((match = cargoPattern.exec(allText)) !== null) {
    cargoMatches.push({ cargo: match[1].toUpperCase(), index: match.index + match[0].length });
  }

  for (var i = 0; i < cargoMatches.length; i++) {
    var cargo = cargoMatches[i].cargo;
    var startIdx = cargoMatches[i].index;
    var endIdx = (i + 1 < cargoMatches.length) ? cargoMatches[i + 1].index - cargoMatches[i + 1].cargo.length - 1 : startIdx + 200;
    var segment = allText.substring(startIdx, Math.min(endIdx, allText.length));

    var mark = '';
    var markMatch = segment.match(/(\d{1,3})\s*[-]\s*(\d{2,5})\s*[-]\s*(\d{3,6}[A-Z]?)/);
    if (markMatch) {
      var cc = markMatch[1];
      while (cc.length < 3) cc = '0' + cc;
      mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
    }

    var sampleSize = '500g';
    var wtMatch = segment.match(/(\d+)\s*(g|lb|lbs)\b/i);
    if (wtMatch) {
      var u = (wtMatch[2] || 'g').toLowerCase();
      sampleSize = u.startsWith('l') ? wtMatch[1] + ' lb' : wtMatch[1] + 'g';
    }

    var desc = '';
    if (markMatch) {
      var afterMark = segment.substring(segment.indexOf(markMatch[0]) + markMatch[0].length);
      afterMark = afterMark.replace(/^\s*\(?'\d{2}\)?\s*/i, '').trim();
      var descMatch = afterMark.match(/^([A-Za-zÀ-ÿ\s#()\-0-9]+?)(?:\d+\s*(?:g|lb))/i);
      if (descMatch) desc = descMatch[1].trim();
    }

    var container = '';
    var contMatch = segment.match(/\b([A-Z]{4}\d{7})\b/);
    if (contMatch) container = contMatch[1];

    var origin = getOriginFromMark(mark);

    orders.push({
      client: 'Osito Coffee',
      sampleOrderNum: '',
      container: container,
      mark: mark,
      cargo: cargo,
      reference: '',
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: fedexAccount,
      bags: '',
      origin: origin,
      description: desc || origin || '',
      comments: ''
    });
  }

  return orders;
}


// ============================================================================
// ROTHFOS (NKG)
// ============================================================================

function extractRothfosOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var reference = '';
  var refMatch = allText.match(/Ctrt\s*#?\s*\/\s*Shipment\s*#?[:\s]*(P\d{5,9})/i);
  if (refMatch) reference = refMatch[1];
  if (!reference) {
    var pMatch = allText.match(/\b(P\d{5,9})\b/);
    if (pMatch) reference = pMatch[1];
  }

  var senderRef = '';
  var srMatch = allText.match(/Sender\s+Sample\s+Ref\s*#?[:\s]*(S\d{5,12})/i);
  if (srMatch) senderRef = srMatch[1];

  var warehouse = detectWarehouse(allText) || 'Continental';

  var sampleSize = '2 lb';
  var sqMatch = allText.match(/Sample\s+Qty[:\s]*(\d+)\s*(Lbs?|lb|g)/i);
  if (sqMatch) sampleSize = sqMatch[1] + ' lb';
  else {
    var emMatch = allText.match(/send\s+(\d+)\s*(lbs?|lb)\b/i);
    if (emMatch) sampleSize = emMatch[1] + ' lb';
  }

  var container = '';
  var contMatch = allText.match(/Container\s*#?[:\s]*([A-Z]{4}\d{7})/i);
  if (contMatch) container = contMatch[1];
  if (!container) {
    var anyM = allText.match(/\b([A-Z]{4}\d{7})\b/);
    if (anyM) container = anyM[1];
  }

  var mark = '';
  var markMatch = allText.match(/Shipment\s+Marks?[:\s]*(\d{1,3})\s*[-]\s*(\d{2,5})\s*[-]\s*(\d{3,6}[A-Z]?)/i);
  if (markMatch) {
    var cc = markMatch[1];
    while (cc.length < 3) cc = '0' + cc;
    mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
  }

  var cargo = '';
  var cargoMatch = allText.match(/Cargo\s+(?:Ref\s+)?#?[:\s]*(C\d{5,7}[A-Z]?)/i);
  if (cargoMatch) cargo = cargoMatch[1].toUpperCase();
  if (!cargo) {
    var cM = allText.match(/\b(C\d{5,7})\b/i);
    if (cM) cargo = cM[1].toUpperCase();
  }

  var bags = '';
  var bagsMatch = allText.match(/Available\s+Qty[:\s]*(\d+)\s+(Kg\s+)?Bags/i);
  if (bagsMatch && !bagsMatch[2]) bags = bagsMatch[1];

  var description = '';
  var origMatch = allText.match(/Origin\s+Grade[^\n]*\n\s*([^\n]+)/i);
  if (origMatch) description = origMatch[1].trim().replace(/\s+[SR]\d{8,}.*$/, '').trim();

  var receiver = '';
  var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
  if (recvMatch) receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  if (!receiver) receiver = 'Rothfos';

  var shipping = extractShipping(allText, emailBody) || 'FedEx 2 Day';
  var origin = getOriginFromMark(mark);

  if (container || mark || cargo) {
    orders.push({
      client: 'Rothfos',
      sampleOrderNum: senderRef || '',
      container: container,
      mark: mark,
      cargo: cargo,
      reference: reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: '',
      bags: bags,
      origin: origin,
      description: description || origin || '',
      comments: ''
    });
  }

  return orders;
}


// ============================================================================
// LOUIS DREYFUS COMPANY (LDC)
// ============================================================================

function extractLdcOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var orderNum = '';
  var orderMatch = allText.match(/Sample\s+Order\s+No\s*[:\s]*(\S+)/i);
  if (orderMatch) orderNum = orderMatch[1];

  var reference = '';
  var refMatch = allText.match(/Our Reference\s+no\.?\s*[:\s]*(S\d{5,12})/i);
  if (refMatch) reference = refMatch[1];
  if (!reference) {
    var sRef = allText.match(/\b(S\d{5,12})\b/);
    if (sRef) reference = sRef[1];
  }

  var warehouse = detectWarehouse(allText) || 'RPM';

  var sampleSize = '2 lb';
  var sqMatch = allText.match(/Sample\s+Quantity\s*[:\s]*(\d+)\s*(lb|lbs|g)/i);
  if (sqMatch) sampleSize = sqMatch[1] + ' lb';

  var receiver = '';
  var recvMatch = allText.match(/Send\s+Sample\s+[Tt]o\s*[:\s]+([^\n]+)/i);
  if (recvMatch) receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  if (!receiver) receiver = 'Louis Dreyfus Company';

  // Origin + Quality for description
  var origin = '';
  var originMatch = allText.match(/Origin\s*[:\s]+([A-Za-z\s]+?)(?=\n|$)/im);
  if (originMatch) origin = originMatch[1].trim();
  var quality = '';
  var qualMatch = allText.match(/Quality\s*[:\s]+([^\n]+)/i);
  if (qualMatch) quality = qualMatch[1].trim();
  var description = [origin, quality].filter(function(x) { return x; }).join(' ');

  var shipping = 'FedEx Standard Overnight';
  if (/standard\s+overnight/i.test(allText)) shipping = 'FedEx Standard Overnight';
  else if (/priority/i.test(allText)) shipping = 'FedEx Priority';
  else shipping = extractShipping(allText, emailBody) || 'FedEx Standard Overnight';

  var fedexAccount = '';
  var fxMatch = allText.match(/Fedex\s+Account[:\s]*(\d[\d\s-]+\d)/i);
  if (fxMatch) fedexAccount = fxMatch[1].replace(/\s+/g, '');

  // Parse table
  var containerPattern = /\b([A-Z]{4}\d{7})\b/g;
  var match;
  var seen = {};

  while ((match = containerPattern.exec(allText)) !== null) {
    var container = match[1];
    if (seen[container]) continue;
    seen[container] = true;

    var lineStart = allText.lastIndexOf('\n', match.index) + 1;
    var lineEnd = allText.indexOf('\n', match.index + container.length);
    var line = allText.substring(lineStart, lineEnd > 0 ? lineEnd : allText.length);

    if (/\bContainer\s+No\b/i.test(line) && /\bMarks\b/i.test(line)) continue;

    // Mark: 003-0279-6076-0 (4-part dash-separated)
    var mark = '';
    var markMatch = line.match(/\b(\d{1,4})\s*[-]\s*(\d{2,5})\s*[-]\s*(\d{3,6}[A-Z]?)(?:\s*[-]\s*\d)?\b/);
    if (markMatch) {
      var cc = markMatch[1];
      while (cc.length < 3) cc = '0' + cc;
      mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
    }

    // Warrant (stored as cargo)
    var cargo = '';
    if (markMatch) {
      var afterMark = line.substring(line.indexOf(markMatch[0]) + markMatch[0].length);
      var warrantMatch = afterMark.match(/\b(\d{5,7})\b/);
      if (warrantMatch) cargo = warrantMatch[1];
    }

    // Bags: last reasonable number on line
    var bags = '';
    var allNums = line.match(/\b(\d{2,4})\b/g);
    if (allNums && allNums.length > 0) {
      for (var i = allNums.length - 1; i >= 0; i--) {
        var n = parseInt(allNums[i]);
        if (n >= 1 && n <= 2000 && allNums[i] !== (markMatch ? markMatch[1] : '')
            && allNums[i] !== (markMatch ? markMatch[2] : '')
            && allNums[i] !== (markMatch ? markMatch[3] : '')) {
          bags = allNums[i];
          break;
        }
      }
    }

    var markOrigin = getOriginFromMark(mark);

    orders.push({
      client: 'Louis Dreyfus Company',
      sampleOrderNum: orderNum,
      container: container,
      mark: mark,
      cargo: cargo,
      reference: reference,
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: fedexAccount,
      bags: bags,
      origin: markOrigin || origin,
      description: description || markOrigin || '',
      comments: ''
    });
  }

  return orders;
}


// ============================================================================
// SUCAFINA NA / TASTIFY
// ============================================================================

function extractSucafinaOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');
  var isTastify = /tastify|notifier@tastify/i.test(allText);

  var warehouse = detectWarehouse(allText) || 'Continental';

  var orderNum = '';
  if (isTastify) {
    var tkMatch = allText.match(/shipment\s+#?(TKM[-]?\d+)/i);
    if (tkMatch) orderNum = tkMatch[1];
  }

  // Receiver
  var receiver = '';
  if (isTastify) {
    var custMatch = allText.match(/Customer[:\s]+([A-Z][A-Z\s&.,']+(?:LLC|INC|CORP|CO|LTD)?)/i);
    if (custMatch) receiver = custMatch[1].trim();
    if (!receiver || receiver.length < 3) {
      var recipMatch = allText.match(/Recipient[:\s]+([A-Za-z][A-Za-z\s]+?)(?=\n|Street)/i);
      if (recipMatch) receiver = recipMatch[1].trim();
    }
    if (!receiver || receiver === '-') {
      var altRecv = allText.match(/recipient\s+information\s+below[^]*?\n\s*([A-Z][A-Za-z\s]+?)\s+\d/i);
      if (altRecv) receiver = altRecv[1].trim();
    }
  } else {
    var recvMatch = allText.match(/SUCAFINA\s+NA\s+(?:INC\.?)?\s*\n/i);
    if (recvMatch) receiver = 'Sucafina NA Inc';
    if (!receiver) {
      var addrMatch = allText.match(/(\d+\s+[A-Za-z\s]+(?:Blvd|Ave|St|Road|Dr))/i);
      if (addrMatch) receiver = 'Sucafina NA Inc';
    }
  }
  if (!receiver) receiver = 'Sucafina NA Inc';

  // Sample size
  var sampleSize = isTastify ? '100g' : '1 lb';
  if (!isTastify) {
    var ssMatch = allText.match(/(\d+)\s*(lb|lbs|g|oz)\s+sample/i);
    if (ssMatch) {
      var u = (ssMatch[2] || 'lb').toLowerCase();
      sampleSize = u.startsWith('g') ? ssMatch[1] + 'g' : ssMatch[1] + ' lb';
    }
  } else {
    var tsMatch = allText.match(/(\d+)\s*g(?:ram)?\b/i);
    if (tsMatch) sampleSize = tsMatch[1] + 'g';
  }

  var shipping = 'FedEx Standard Overnight';
  if (/overnight/i.test(allText)) shipping = 'FedEx Standard Overnight';
  else if (/2nd?\s*day/i.test(allText)) shipping = 'FedEx 2 Day';
  else if (/ground/i.test(allText)) shipping = 'FedEx Ground';
  else shipping = extractShipping(allText, emailBody) || 'FedEx Standard Overnight';

  // Find container+cargo pairs as anchors
  var lines = allText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    if (/\bContainer\s*#?\b/i.test(line) && /\bCargo\s*#?\b/i.test(line)) continue;
    if (/\bWHSE\b/i.test(line) && /\bOrigin\b/i.test(line) && /\bMarks\b/i.test(line)) continue;

    var contCargoPattern = /\b([A-Z]{4}\d{7})\s*(C\d{5,7})\b/gi;
    var ccMatch;

    while ((ccMatch = contCargoPattern.exec(line)) !== null) {
      var container = ccMatch[1];
      var cargo = ccMatch[2].toUpperCase();

      var beforeText = line.substring(0, ccMatch.index);
      var afterCargo = line.substring(ccMatch.index + ccMatch[0].length);

      // ICO Mark
      var mark = '';
      var allMarks = [];
      var markRe = /\b(\d{1,3})\/(\d{2,5})\/(\d{3,6}[A-Z]?)\b/g;
      var mm;
      while ((mm = markRe.exec(beforeText)) !== null) {
        if (/^20[0-3]\d$/.test(mm[1])) continue;
        var cc = mm[1];
        while (cc.length < 3) cc = '0' + cc;
        allMarks.push({ mark: cc + '/' + mm[2] + '/' + mm[3], endIndex: mm.index + mm[0].length });
      }
      if (allMarks.length > 0) mark = allMarks[allMarks.length - 1].mark;

      // Bags
      var bags = '';
      if (!isTastify) {
        var bagsAfter = afterCargo.match(/\b(\d{1,4})\b/);
        if (bagsAfter && parseInt(bagsAfter[1]) <= 2000) bags = bagsAfter[1];
        if (!bags) {
          var beforeCont = beforeText;
          if (allMarks.length > 0) beforeCont = beforeText.substring(allMarks[allMarks.length - 1].endIndex);
          var bagsBefore = beforeCont.match(/\b(\d{1,4})\b/);
          if (bagsBefore && parseInt(bagsBefore[1]) <= 2000) bags = bagsBefore[1];
        }
      } else {
        if (allMarks.length > 0) {
          var afterMarkText = beforeText.substring(allMarks[allMarks.length - 1].endIndex);
          var tastifyBags = afterMarkText.match(/^(\d{1,4})(?:No|Yes)/i);
          if (tastifyBags) bags = tastifyBags[1];
        }
      }

      // NP Reference
      var npRef = '';
      var npMatch = beforeText.match(/\b(NP[A-Z]{2}-\d{4,6}(?:-\d)?)\b/i);
      if (npMatch) npRef = npMatch[1].toUpperCase();

      // Description
      var description = '';
      var originWords = beforeText.match(/\b(COLO|Colombia|Brazil|Peru|Honduras|Guatemala|Costa Rica|Mexico|Ethiopia|Kenya|Rwanda|Burundi|Indonesia|Vietnam|India|Uganda|Tanzania|Nicaragua|El Salvador|Ecuador|Bolivia)\b/i);
      if (originWords) {
        var origEnd = beforeText.indexOf(originWords[0]) + originWords[0].length;
        var gradeText = beforeText.substring(origEnd).trim().replace(/\s*NP[A-Z]{2}[-\d]+.*$/i, '').trim();
        description = gradeText.length > 2 ? originWords[1] + ' ' + gradeText : originWords[1];
      }
      description = description.replace(/\bCOLO\b/i, 'Colombia');
      if (!description) description = getOriginFromMark(mark) || '';

      orders.push({
        client: isTastify ? 'Sucafina NA (Tastify)' : 'Sucafina NA',
        sampleOrderNum: orderNum,
        container: container,
        mark: mark,
        cargo: cargo,
        reference: npRef,
        warehouse: warehouse,
        receiver: receiver,
        sampleSize: sampleSize,
        shipping: shipping,
        fedexAccount: '',
        bags: bags,
        origin: getOriginFromMark(mark),
        description: description,
        comments: ''
      });
    }
  }

  // Fallback: containers and cargos separately
  if (orders.length === 0) {
    var containers = [];
    var cargos = [];
    var marks = [];
    var contRe = /\b([A-Z]{4}\d{7})\b/g;
    var cargRe = /\b(C\d{5,7})\b/gi;
    var markRe2 = /\b(\d{1,3})\/(\d{2,5})\/(\d{3,6}[A-Z]?)\b/g;
    var m;
    while ((m = contRe.exec(allText)) !== null) containers.push(m[1]);
    while ((m = cargRe.exec(allText)) !== null) cargos.push(m[1].toUpperCase());
    while ((m = markRe2.exec(allText)) !== null) {
      if (/^20[0-3]\d$/.test(m[1])) continue;
      var mcc = m[1]; while (mcc.length < 3) mcc = '0' + mcc;
      marks.push(mcc + '/' + m[2] + '/' + m[3]);
    }

    var count = Math.max(containers.length, cargos.length, marks.length);
    for (var j = 0; j < count; j++) {
      var mk = marks[j] || marks[0] || '';
      orders.push({
        client: isTastify ? 'Sucafina NA (Tastify)' : 'Sucafina NA',
        sampleOrderNum: orderNum,
        container: containers[j] || '',
        mark: mk,
        cargo: cargos[j] || '',
        reference: '',
        warehouse: warehouse,
        receiver: receiver,
        sampleSize: sampleSize,
        shipping: shipping,
        fedexAccount: '',
        bags: '',
        origin: getOriginFromMark(mk),
        description: getOriginFromMark(mk) || '',
        comments: ''
      });
    }
  }

  return orders;
}


// ============================================================================
// THE COFFEE SOURCE LLC
// ============================================================================

function extractCoffeeSourceOrders(emailBody, pdfText, subject) {
  var orders = [];
  var allText = (emailBody || '') + '\n' + (pdfText || '');

  var reference = '';
  var refMatch = allText.match(/Ctrt\s*#?\s*\/\s*Shipment\s*#?[:\s]*(P\d{5,9})/i);
  if (refMatch) reference = refMatch[1];

  var senderRef = '';
  var srMatch = allText.match(/Sender\s+Sample\s+Ref\s*#?[:\s]*([^\n]+?)(?=\s*Client|\s*$)/i);
  if (srMatch) senderRef = srMatch[1].trim();

  var clientRef = '';
  var crMatch = allText.match(/Client\s+Contract\s+Ref\s*#?[:\s]*([^\n]+)/i);
  if (crMatch) clientRef = crMatch[1].trim();

  var warehouse = detectWarehouse(allText) || 'Continental';

  var sampleSize = '1 lb';
  var sqMatch = allText.match(/Sample\s+Qty[:\s]*(\d+)\s*(Lbs?|lb|g)/i);
  if (sqMatch) sampleSize = sqMatch[1] + ' lb';

  var container = '';
  var contMatch = allText.match(/Container\s*#?[:\s]*([A-Z]{4}\d{7})/i);
  if (contMatch) container = contMatch[1];
  if (!container) {
    var anyM = allText.match(/\b([A-Z]{4}\d{7})\b/);
    if (anyM) container = anyM[1];
  }

  // Shipment Marks (2-digit country code OK)
  var mark = '';
  var markMatch = allText.match(/Shipment\s+Marks?[:\s]*(\d{1,3})\s*[-]\s*(\d{2,5})\s*[-]\s*(\d{3,6}[A-Z]?)/i);
  if (markMatch) {
    var cc = markMatch[1];
    while (cc.length < 3) cc = '0' + cc;
    mark = cc + '/' + markMatch[2] + '/' + markMatch[3];
  }
  if (!mark) {
    var slashMark = allText.match(/\b(\d{1,4})\s*\/\s*(\d{2,5})\s*\/\s*(\d{3,6}[A-Z]?)\b/);
    if (slashMark && !/^20[0-3]\d$/.test(slashMark[1])) {
      var cc2 = slashMark[1];
      while (cc2.length < 3) cc2 = '0' + cc2;
      mark = cc2 + '/' + slashMark[2] + '/' + slashMark[3];
    }
  }

  var cargo = '';
  var cargoMatch = allText.match(/Cargo\s+(?:Ref\s+)?#?[:\s]*(C\d{5,7}[A-Z]?)/i);
  if (cargoMatch) cargo = cargoMatch[1].toUpperCase();
  if (!cargo) {
    var cM = allText.match(/\b(C\d{5,7})\b/i);
    if (cM) cargo = cM[1].toUpperCase();
  }

  var bags = '';
  var bagsMatch = allText.match(/Available\s+Qty[:\s]*(\d+)\s+(Kg\s+)?Bags/i);
  if (bagsMatch && !bagsMatch[2]) bags = bagsMatch[1];
  if (!bags) {
    var parBags = allText.match(/\b(\d{2,4})\s+BAGS\b/i);
    if (parBags) bags = parBags[1];
  }

  var description = '';
  var origMatch = allText.match(/Origin\s+Grade[^\n]*\n\s*([^\n]+)/i);
  if (origMatch) description = origMatch[1].trim().replace(/\s+[SR]\d{8,}.*$/, '').trim();

  var receiver = '';
  var recvMatch = allText.match(/Send\s+Samples?\s+to[:\s]+([^\n]+)/i);
  if (recvMatch) receiver = recvMatch[1].trim().replace(/\d{3,}.*$/, '').replace(/,\s*$/, '').trim();
  if (!receiver) receiver = 'The Coffee Source LLC';

  var shipping = extractShipping(allText, emailBody) || 'FedEx Standard Overnight';

  // FedEx account from courier table
  var fedexAccount = '';
  var fxMatch = allText.match(/Fedex\s+Courier\s*\/?\s*Account\s*\n?\s*(\d[\d-]+)/i);
  if (fxMatch) fedexAccount = fxMatch[1];
  if (!fedexAccount) {
    var fxMatch2 = allText.match(/(?:Fedex|FedEx)\s+(?:Account|Acct)[:\s#]*(\d[\d-]+)/i);
    if (fxMatch2) fedexAccount = fxMatch2[1];
  }

  var origin = getOriginFromMark(mark);

  if (container || mark || cargo) {
    orders.push({
      client: 'The Coffee Source LLC',
      sampleOrderNum: senderRef || '',
      container: container,
      mark: mark,
      cargo: cargo,
      reference: reference || clientRef || '',
      warehouse: warehouse,
      receiver: receiver,
      sampleSize: sampleSize,
      shipping: shipping,
      fedexAccount: fedexAccount,
      bags: bags,
      origin: origin,
      description: description || origin || '',
      comments: ''
    });
  }

  return orders;
}
