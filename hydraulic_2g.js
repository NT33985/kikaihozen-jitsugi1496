/* 機械保全技能検定 実技・2級「油圧回路図」 — Excel 2026-11-09 準拠 */
(() => {
  'use strict';
  const CAT='hydraulic_diagram';
  const MODE_STRUCTURE='構造図→名称';
  const MODE_SYMBOL='名称→図記号';
  const MODE_FUNCTION='図記号→機能';
  const MODE_ALL='すべて出題';
  const MODES=[MODE_STRUCTURE,MODE_SYMBOL,MODE_FUNCTION];
  // row, D列 名称, F列 機能（原文）, 4択表示用の短文。
  // 図と名称と機能の正誤判定はExcelの列ごとの組合せに準拠。
  const ROWS=[
    [8,'定容量形ベーンポンプ','電動モータなどの動力で回転を得て、作動油を圧油として油圧回路に供給する。ポンプ1回転当たりの吐出量が可変できない','一定の吐出量で作動油を送る'],
    [9,'電磁切換弁付きリリーフ弁','最高使用圧力の設定と電気信号により、ポンプを無負荷運転にする','最高圧力を設定し無負荷運転にする'],
    [10,'カウンタバランス弁','シリンダ動作に背圧をかけて動作を安定させ、リリーフ弁（ブレーキ動作）と同じ働きをする','背圧によりシリンダの動作を安定させる'],
    [11,'流量調整弁（圧力補償付き）','タンクへ戻す油量調整により、シリンダースピードを設定する','戻り油量を調整してシリンダ速度を設定'],
    [12,'パイロット操作切換弁','パイロット圧力によりスプールが操作され、作動油の流れの方向を切り換える','パイロット圧力で流れの方向を切り換える'],
    [13,'機械操作切換弁','機械的方法によりアクチュエータの運動方向を切り換える','機械的操作で運動方向を切り換える'],
    [14,'両ロッド油圧シリンダー','流体エネルギー（流量、圧力）を直線往復運動に変換して仕事動作を行う','油圧を直線往復運動に変換する'],
    [15,'圧力計','回路内圧力の表示を行う','油圧回路内の圧力を表示する'],
    [16,'サクションストレーナー','作動油に侵入している固形粒子やゴミを除去し、回路内に持ち込ませない','作動油中のゴミや固形粒子を除去する'],
    [17,'エアブリーザー','給油口とエレメントの組み合わせでタンク内の空気の出入りを可能とし、かつ粉塵などの侵入を防ぐ','タンクの通気を確保し粉塵侵入を防ぐ']
  ];
  if(typeof SETTINGS==='undefined'||typeof DB==='undefined'||!SETTINGS['2級']||typeof getCards!=='function'){
    console.error('[油圧回路図] 既存実技アプリの初期化を確認できません');
    return;
  }
  // 2級のみ追加。既存の分野・モード・進捗情報を上書きしない。
  SETTINGS['2級'][CAT]={name:'油圧回路図',adopted:[...MODES,MODE_ALL],hold:[]};
  DB.categories.push({id:CAT,name:'油圧回路図',levels:['2級'],answer_modes:MODES});
  let loaded=false;
  const prepared=ROWS.map(([row,name,full,short],i)=>{
    const id=`hydraulic_diagram_${String(i+1).padStart(2,'0')}`;
    const picStructure=`${id}_structure`,picSymbol=`${id}_symbol`;
    const item={id,category_id:CAT,levels:['2級'],active:true,answers:{'名称':name,'機能':full},
      excel_row:row,short_function:short,image_ids:[picStructure,picSymbol]};
    DB.items.push(item);
    for(const [picId,kind] of [[picStructure,'構造図'],[picSymbol,'油圧用図記号']]){
      const entry={image_id:picId,item_id:id,kind,data:'',source:'油圧回路図.xlsx',source_page:row};
      DB.photos.push(entry);imgs[picId]=entry;
    }
    return item;
  });
  group[CAT]=prepared;
  cats[CAT]=DB.categories.at(-1);
  const previousGetCards=getCards;
  getCards=function(category,mode){
    if(category!==CAT)return previousGetCards(category,mode);
    if(grade!=='2級'||!loaded)return [];
    const modes=mode===MODE_ALL?MODES:[mode];
    return modes.filter(m=>MODES.includes(m)).flatMap(m=>prepared.map(item=>({
      item,mode:m,pic:imgs[item.image_ids[m===MODE_STRUCTURE?0:1]]
    })));
  };

  const css=document.createElement('style');
  css.textContent=`
  .hydraulic-name-text{font-size:clamp(21px,3vw,32px);font-weight:800;text-align:center;line-height:1.55;color:#140a9a;overflow-wrap:anywhere;}
  .hydraulic-name-text small{display:block;font-size:12px;color:#65738c;font-weight:700;margin-bottom:12px}
  .hydraulic-choice-wrap{position:relative;min-width:0}
  .hydraulic-choice-wrap>.answer{height:100%;width:100%}
  .hydraulic-choice-wrap>.answer.photoChoice{padding-right:34px}
  .hydraulic-choice-zoom{position:absolute;right:6px;top:6px;width:30px;height:30px;z-index:2;display:grid;place-items:center;background:#f0f3ff;border:1px solid #c8d0f3;color:#140a9a;border-radius:7px;font-size:15px;cursor:zoom-in}
  .hydraulic-choice-zoom:focus-visible{outline:3px solid #3668d8;outline-offset:2px}
  .hydraulic-label{font-size:clamp(11px,2.6vw,14px);line-height:1.36}
  body.in-quiz .result.hydraulic-result{overflow-y:auto;max-height:min(180px,24dvh)}
  @media(max-width:720px){body.in-quiz .result.hydraulic-result{max-height:min(140px,23dvh)}.hydraulic-choice-zoom{width:26px;height:26px}}
  `;
  document.head.append(css);

  const previousShowQuestion=showQuestion;
  showQuestion=function(){
    if(catSel.value!==CAT)return previousShowQuestion();
    document.body.classList.remove('mode-overview');
    answersEl.replaceChildren();nextBtn.disabled=true;done=false;
    resultEl.className='result';resultEl.replaceChildren();qtag.textContent='油圧回路図';
    if(!deck.length){photoEl.hidden=true;photoName.hidden=true;emptyPhoto.hidden=false;emptyPhoto.textContent='出題データの読み込みに失敗しました。';return;}
    if(idx>=deck.length){deck=shuffle(cards);idx=0;}
    current=deck[idx];
    const item=current.item,mode=current.mode;
    const isChoiceImage=mode===MODE_SYMBOL;
    const isStructure=mode===MODE_STRUCTURE;
    photoName.hidden=true;photoNameText.textContent='';
    photoEl.hidden=isChoiceImage;
    emptyPhoto.hidden=!isChoiceImage;
    emptyPhoto.replaceChildren();
    if(isChoiceImage){
      const display=document.createElement('div');display.className='hydraulic-name-text';
      const lead=document.createElement('small');lead.textContent='この名称に対応する図記号は？';
      display.append(lead,document.createTextNode(item.answers['名称']));emptyPhoto.append(display);
      $('imageInfo').textContent='図記号は選択肢の 🔍 から拡大できます';
    }else{
      photoEl.src=current.pic.data;
      photoEl.alt=(isStructure?'構造図':'油圧用図記号')+'：タップして拡大';
      $('imageInfo').textContent=(isStructure?'構造図':'油圧用図記号')+'／タップで拡大';
    }
    position.textContent=`${idx+1} / ${deck.length}問`;
    promptEl.textContent=isStructure?'この構造図の機器名称は？':isChoiceImage?'正しい油圧用図記号を選んでください':'この図記号が表す機能は？';
    const others=prepared.filter(i=>i.id!==item.id);
    const options=shuffle([item,...shuffle(others).slice(0,3)]);
    for(const [i,option] of options.entries()){
      const btn=document.createElement('button');btn.type='button';
      btn.className='answer'+(isChoiceImage?' photoChoice':'');btn.dataset.itemid=option.id;
      const badge=document.createElement('span');badge.className='circle';badge.textContent=String.fromCharCode(65+i);btn.append(badge);
      if(isChoiceImage){
        const im=document.createElement('img');im.src=imgs[option.image_ids[1]].data;
        im.alt=`選択肢 ${String.fromCharCode(65+i)} の図記号`;btn.append(im);
        const wrap=document.createElement('div');wrap.className='hydraulic-choice-wrap';
        const zoomButton=document.createElement('button');zoomButton.type='button';
        zoomButton.className='hydraulic-choice-zoom';zoomButton.textContent='🔍';
        zoomButton.setAttribute('aria-label',`選択肢 ${String.fromCharCode(65+i)} の図記号を拡大`);
        zoomButton.onclick=()=>{zoomImg.src=im.src;$('zoomName').hidden=true;zoom.classList.add('visible');zoomClose.focus();};
        btn.onclick=()=>hydraulicChoose(option.id,item.id);
        wrap.append(btn,zoomButton);answersEl.append(wrap);
      }else{
        const name=document.createElement('span');name.className='hydraulic-label';
        name.textContent=isStructure?option.answers['名称']:option.short_function;
        btn.append(name);btn.onclick=()=>hydraulicChoose(option.id,item.id);answersEl.append(btn);
      }
    }
    updateSession();
  };
  function hydraulicChoose(pick,answerId){
    if(done)return;done=true;
    const item=current.item,mode=current.mode,correct=pick===answerId;
    answerCount++;history.answered++;history.streak=correct?history.streak+1:0;
    if(correct){correctCount++;history.correct++;}
    history.best=Math.max(history.best,history.streak);saveHistory();
    Challenge.record({correct,valid:true,grade:'2級',category:CAT,mode,itemId:item.id});
    for(const b of answersEl.querySelectorAll('.answer')){
      b.disabled=true;
      if(b.dataset.itemid===answerId)b.classList.add('correct');
      else if(b.dataset.itemid===pick)b.classList.add('incorrect');
    }
    for(const zoomButton of answersEl.querySelectorAll('.hydraulic-choice-zoom'))zoomButton.disabled=false;
    const label=item.answers['名称'];
    // 正解後は、短縮した選択肢ではなくExcel F列の説明を原文表示する。
    showResult(correct?'success':'failure',correct?'🎉 正解！':'不正解。解説を確認しよう',label,`Excelの機能説明：${item.answers['機能']}`);
    resultEl.classList.add('hydraulic-result');
    nextBtn.disabled=false;updateSession();
  }
  const sprite=new Image();
  sprite.onload=()=>{
    try{
      const canvas=document.createElement('canvas');canvas.width=400;canvas.height=220;
      const ctx=canvas.getContext('2d',{willReadFrequently:false});
      if(!ctx)throw new Error('Canvas非対応');
      for(const [i,item] of prepared.entries()){
        for(let col=0;col<2;col++){
          ctx.clearRect(0,0,400,220);ctx.fillStyle='#fff';ctx.fillRect(0,0,400,220);
          ctx.drawImage(sprite,col*400,i*220,400,220,0,0,400,220);
          imgs[item.image_ids[col]].data=canvas.toDataURL('image/webp',0.88);
        }
      }
      loaded=true;
      populateCategories();
    }catch(error){console.error('[油圧回路図] 画像展開エラー',error);}
  };
  sprite.onerror=()=>console.error('[油圧回路図] 画像ファイルを読み込めません');
  sprite.src='hydraulics_sprite.webp?v=20261109';
  populateCategories();
})();