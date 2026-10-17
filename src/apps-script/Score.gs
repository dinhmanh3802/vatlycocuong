/** CHẤM ĐIỂM — dùng khi sửa câu trả lời của một em, và để đối chiếu với điểm TNMaker */

function round2(x) { return Math.round(x * 100 + 1e-9) / 100; }

function normP3(s) {
  return String(s == null ? '' : s).trim().replace(/\./g, ',').replace(/_+$/, '');
}
function normP2(s) {
  var t = String(s == null ? '' : s), out = '';
  for (var i = 0; i < 4; i++) {
    var c = (t.charAt(i) || '_').toUpperCase();
    out += (c === 'Đ' || c === 'S') ? c : '_';
  }
  return out;
}

/**
 * cfg: cấu hình bài {p1:{on,soCau,diem}, p2:{on,soCau,d1..d4}, p3:{on,soCau,diem}}
 * tl:  câu trả lời của em {p1:['A',...], p2:['ĐSSĐ',...], p3:['8','0,31',...]}
 * key: đáp án đúng cùng dạng
 * Trả về điểm từng phần, tổng, và mảng đúng/sai từng câu (null nếu thiếu đáp án).
 */
function chamBai(cfg, tl, key) {
  tl = tl || {}; key = key || {};
  var out = { p1: 0, p2: 0, p3: 0, tong: 0, thieuKey: false, dung: { p1: [], p2: [], p3: [] } };

  if (cfg.p1 && cfg.p1.on) {
    var n1 = Number(cfg.p1.soCau) || 0, per1 = n1 ? Number(cfg.p1.diem) / n1 : 0;
    for (var i = 0; i < n1; i++) {
      var k1 = key.p1 && key.p1[i] ? String(key.p1[i]).trim().toUpperCase() : '';
      if (!k1) { out.thieuKey = true; out.dung.p1.push(null); continue; }
      var a1 = tl.p1 && tl.p1[i] ? String(tl.p1[i]).trim().toUpperCase() : '';
      var ok1 = a1 === k1;
      out.dung.p1.push(ok1 ? 1 : 0);
      if (ok1) out.p1 += per1;
    }
  }

  if (cfg.p2 && cfg.p2.on) {
    var n2 = Number(cfg.p2.soCau) || 0;
    var pts = [0, Number(cfg.p2.d1) || 0, Number(cfg.p2.d2) || 0, Number(cfg.p2.d3) || 0, Number(cfg.p2.d4) || 0];
    for (var j = 0; j < n2; j++) {
      var k2 = key.p2 && key.p2[j] ? normP2(key.p2[j]) : '';
      if (!k2 || k2.indexOf('_') !== -1) { out.thieuKey = true; out.dung.p2.push(null); continue; }
      var a2 = normP2(tl.p2 && tl.p2[j]);
      var cnt = 0;
      for (var m = 0; m < 4; m++) if (a2.charAt(m) === k2.charAt(m)) cnt++;
      out.dung.p2.push(cnt);
      out.p2 += pts[cnt];
    }
  }

  if (cfg.p3 && cfg.p3.on) {
    var n3 = Number(cfg.p3.soCau) || 0, per3 = n3 ? Number(cfg.p3.diem) / n3 : 0;
    for (var q = 0; q < n3; q++) {
      var k3 = key.p3 ? normP3(key.p3[q]) : '';
      if (!k3) { out.thieuKey = true; out.dung.p3.push(null); continue; }
      var a3 = normP3(tl.p3 && tl.p3[q]);
      var ok3 = a3 !== '' && a3 === k3;
      out.dung.p3.push(ok3 ? 1 : 0);
      if (ok3) out.p3 += per3;
    }
  }

  out.p1 = round2(out.p1); out.p2 = round2(out.p2); out.p3 = round2(out.p3);
  out.tong = round2(out.p1 + out.p2 + out.p3);
  return out;
}
