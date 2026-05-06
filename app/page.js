"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

const ALLERGEN_RULES = [
  { allergen: "Tree Nuts", keywords: ["almond","walnut","pecan","cashew","hazelnut","pistachio","macadamia"] },
  { allergen: "Peanuts", keywords: ["peanut"] },
  { allergen: "Seeds", keywords: ["seed","seeds","sunflower seed","pumpkin seed","chia","flax","hemp seed"] },
  { allergen: "Soy", keywords: ["soy","tofu","edamame","miso","tamari"] },
  { allergen: "Gluten", keywords: ["wheat","flour","gluten","bread","pasta","semolina","barley","rye","panko"] },
  { allergen: "Milk / Dairy", keywords: ["milk","cream","butter","cheese","yogurt","parmesan","mozzarella","ricotta","cream cheese"] },
  { allergen: "Egg", keywords: ["egg","eggs","mayonnaise","aioli"], exclude: ["eggplant"] },
  { allergen: "Fish", keywords: ["salmon","tuna","cod","anchovy","fish","sardine"] },

  // ✅ FIXED SHELLFISH
  {
    allergen: "Shellfish",
    keywords: ["shrimp","crab","lobster","mussel","oyster","scallop"],
    exclude: ["clam shell","clamshell","packed in","pk","tray","case"]
  },

  { allergen: "Sesame", keywords: ["sesame","tahini"] },
  { allergen: "Mustard", keywords: ["mustard"] }
];

const cleanText = (v) =>
  String(v || "").toUpperCase().replace(/\s+/g," ").trim();

const normalizeVenue = (v) =>
  cleanText(v)
    .replace(/\s*-\s*VV$/g,"")
    .replace(/\s*VV$/g,"")
    .trim();

export default function App() {
  const [consumptionRows,setConsumptionRows]=useState([]);
  const [recipeRows,setRecipeRows]=useState([]);
  const [products,setProducts]=useState([]);
  const [selectedProduct,setSelectedProduct]=useState("");
  const [selectedRecipe,setSelectedRecipe]=useState(null);
  const [search,setSearch]=useState("");
  const [userShip,setUserShip]=useState("");
  const [loggedIn,setLoggedIn]=useState(false);

  const shipColumns={BRL:8,RL:11,SC:14,VL:17};

  const readExcel=(file,cb)=>{
    const r=new FileReader();
    r.onload=(e)=>{
      const wb=XLSX.read(e.target.result,{type:"binary"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      cb(XLSX.utils.sheet_to_json(ws,{header:1}));
    };
    r.readAsBinaryString(file);
  };

  const uploadConsumption=(e)=>{
    const f=e.target.files?.[0];
    if(!f)return;
    readExcel(f,(rows)=>{
      setConsumptionRows(rows);
      setProducts([...new Set(rows.slice(1).map(r=>r[6]).filter(Boolean))]);
    });
  };

  const uploadRecipe=(e)=>{
    const f=e.target.files?.[0];
    if(!f)return;
    readExcel(f,setRecipeRows);
  };

  const consumptionData=useMemo(()=>consumptionRows.slice(1),[consumptionRows]);
  const recipeData=useMemo(()=>recipeRows.slice(1),[recipeRows]);

  const productMatches=(p,row)=>{
    const sel=cleanText(p);
    const a=cleanText(row[12]);
    const n=cleanText(row[7]);

    if(a===sel||n===sel)return true;
    if(a.length>12&&(sel.includes(a)||a.includes(sel)))return true;
    if(n.length>12&&(sel.includes(n)||n.includes(sel)))return true;
    return false;
  };

  const getConsumption=(p)=>{
    let v="";
    const res={};
    consumptionData.forEach(r=>{
      if(r[2])v=r[2];
      const key=normalizeVenue(v);
      if(r[6]!==p)return;
      if(!res[key])res[key]={display:v,ships:{}};
      SHIPS.forEach(s=>{
        res[key].ships[s]=(res[key].ships[s]||0)+(Number(r[shipColumns[s]])||0);
      });
    });
    return res;
  };

  const getRequired=(p)=>{
    const req={};
    recipeData.forEach(r=>{
      if(!productMatches(p,r))return;
      const v=normalizeVenue(r[1]);
      if(!req[v])req[v]={display:r[1]};
    });
    return req;
  };

  const combine=(p)=>{
    const a=getConsumption(p);
    const r=getRequired(p);
    const keys=[...new Set([...Object.keys(a),...Object.keys(r)])];

    return keys.map(k=>{
      const ships={};
      SHIPS.forEach(s=>ships[s]=a[k]?.ships?.[s]||0);

      return {
        name:a[k]?.display||r[k]?.display,
        ships,
        missing:SHIPS.filter(s=>r[k]&&(ships[s]===0))
      };
    });
  };

  const getRecipes=(p)=>{
    const res={};
    recipeData.forEach(r=>{
      if(!productMatches(p,r))return;
      const code=r[15]||"N/A";
      const name=r[16]||"Unnamed";
      res[code+name]={code,name};
    });
    return Object.values(res);
  };

  const getIngredients=(rec)=>{
    const items=[];
    recipeData.forEach(r=>{
      if(r[15]===rec.code&&r[16]===rec.name){
        const p=r[12]||r[7];
        if(p)items.push(p);
      }
    });
    return [...new Set(items)];
  };

  const getSub=(name)=>{
    const res=[];
    recipeData.forEach(r=>{
      if(cleanText(r[16])===cleanText(name)){
        const p=r[12]||r[7];
        if(p&&cleanText(p)!==cleanText(name))res.push(p);
      }
    });
    return [...new Set(res)];
  };

  const detectAllergens=(items)=>{
    const found={};

    const check=(txt,label)=>{
      const lower=txt.toLowerCase();
      ALLERGEN_RULES.forEach(rule=>{
        const excluded=rule.exclude?.some(e=>lower.includes(e));
        const hit=!excluded && rule.keywords.find(k=>lower.includes(k));
        if(hit){
          if(!found[rule.allergen])found[rule.allergen]=new Set();
          found[rule.allergen].add(label);
        }
      });
    };

    items.forEach(p=>{
      check(p,p);
      getSub(p).forEach(s=>check(s,p+" → "+s));
    });

    return Object.entries(found);
  };

  if(!loggedIn){
    return (
      <div style={{padding:40}}>
        <img src="/virgin-logo.png" style={{height:60}}/>
        <select onChange={e=>setUserShip(e.target.value)}>
          <option>Select ship</option>
          {SHIPS.map(s=><option key={s}>{s}</option>)}
        </select>
        <button onClick={()=>setLoggedIn(true)}>Enter</button>
      </div>
    );
  }

  const venues=selectedProduct?combine(selectedProduct):[];
  const recipes=selectedProduct?getRecipes(selectedProduct):[];
  const ingredients=selectedRecipe?getIngredients(selectedRecipe):[];
  const allergens=selectedRecipe?detectAllergens(ingredients):[];

  return (
    <div style={{padding:20}}>
      <img src="/virgin-logo.png" style={{height:50}}/>

      <input type="file" onChange={uploadConsumption}/>
      <input type="file" onChange={uploadRecipe}/>

      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search"/>

      {products.filter(p=>p.toLowerCase().includes(search.toLowerCase()))
        .map(p=><div key={p} onClick={()=>setSelectedProduct(p)}>{p}</div>)}

      {selectedProduct && (
        <>
          <h2>{selectedProduct}</h2>

          {venues.map((v,i)=>(
            <div key={i}>
              <b>{v.name}</b>
              {SHIPS.map(s=>(
                <span key={s}
                  style={{color:v.missing.includes(s)?"red":"black",margin:5}}>
                  {s}:{v.ships[s]}
                </span>
              ))}
            </div>
          ))}

          <h3>Recipes</h3>
          {recipes.map(r=>(
            <div key={r.code} onClick={()=>setSelectedRecipe(r)}>
              {r.name}
            </div>
          ))}

          {selectedRecipe && (
            <>
              <h3>Ingredients</h3>
              {ingredients.map(i=>(
                <div key={i}>
                  {i}
                  {getSub(i).map(s=>(
                    <div key={s} style={{marginLeft:20}}>{s}</div>
                  ))}
                </div>
              ))}

              <h3>Allergens</h3>
              {allergens.map(([a,items])=>(
                <div key={a}>
                  <b>{a}</b>
                  <ul>{[...items].map(i=><li key={i}>{i}</li>)}</ul>
                </div>
              ))}
            </>
          )}
        </>
      )}
    </div>
  );
}
