"""
pptx_animations.py — Animation helper module for python-pptx

Injects native Microsoft PowerPoint OpenXML entrance animations into slides:
- Sequential on-click reveals for bounding boxes (cards).
- Progressive on-click disclosure for individual bullet points / paragraphs.
"""

from pptx.oxml import parse_xml


def generate_timing_element(steps, bld_shapes=None):
    """
    Builds the OpenXML <p:timing> element with sequential entrance effects.
    
    steps: list of clicks. Each click is a list of target dicts:
        target = {
            'spid': int,
            'p_range': (start, end) or None  # None for whole shape
        }
    bld_shapes: set/list of shape IDs that have paragraph builds
    """
    id_counter = 3
    seq_children_xml = []
    
    for step_targets in steps:
        if not step_targets:
            continue
            
        par_effects = []
        for idx, tgt in enumerate(step_targets):
            spid = tgt['spid']
            p_range = tgt.get('p_range')
            node_type = "clickEffect" if idx == 0 else "withEffect"
            
            c_tn_id = id_counter
            id_counter += 1
            bhvr_id = id_counter
            id_counter += 1
            
            if p_range is not None:
                st, end = p_range
                tgt_el = f"""<p:tgtEl>
                              <p:spTgt spid="{spid}">
                                <p:txEl>
                                  <p:pRg st="{st}" end="{end}"/>
                                </p:txEl>
                              </p:spTgt>
                            </p:tgtEl>"""
            else:
                tgt_el = f"""<p:tgtEl>
                              <p:spTgt spid="{spid}"/>
                            </p:tgtEl>"""
            
            effect_xml = f"""
            <p:par>
              <p:cTn id="{c_tn_id}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node_type}">
                <p:stCondLst>
                  <p:cond delay="0"/>
                </p:stCondLst>
                <p:childTnLst>
                  <p:set>
                    <p:cBhvr>
                      <p:cTn id="{bhvr_id}" dur="1" fill="hold">
                        <p:stCondLst>
                          <p:cond delay="0"/>
                        </p:stCondLst>
                      </p:cTn>
                      {tgt_el}
                      <p:attrNameLst>
                        <p:attrName>style.visibility</p:attrName>
                      </p:attrNameLst>
                    </p:cBhvr>
                    <p:to>
                      <p:strVal val="visible"/>
                    </p:to>
                  </p:set>
                </p:childTnLst>
              </p:cTn>
            </p:par>"""
            par_effects.append(effect_xml)
            
        wrapper_id1 = id_counter
        id_counter += 1
        wrapper_id2 = id_counter
        id_counter += 1
        
        step_xml = f"""
        <p:par>
          <p:cTn id="{wrapper_id1}" fill="hold">
            <p:stCondLst>
              <p:cond delay="indefinite"/>
            </p:stCondLst>
            <p:childTnLst>
              <p:par>
                <p:cTn id="{wrapper_id2}" fill="hold">
                  <p:stCondLst>
                    <p:cond delay="0"/>
                  </p:stCondLst>
                  <p:childTnLst>
                    {"".join(par_effects)}
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
        </p:par>"""
        seq_children_xml.append(step_xml)
        
    bld_p_xml = ""
    if bld_shapes:
        bld_p_elements = "".join([f'<p:bldP spid="{spid}" grpId="0" uiExpand="1" build="p"/>' for spid in sorted(bld_shapes)])
        bld_p_xml = f"<p:bldLst>{bld_p_elements}</p:bldLst>"
        
    timing_xml = f"""
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    {"".join(seq_children_xml)}
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst>
                  <p:cond evt="onPrev" delay="0">
                    <p:tgtEl>
                      <p:sldTgt/>
                    </p:tgtEl>
                  </p:cond>
                </p:prevCondLst>
                <p:nextCondLst>
                  <p:cond evt="onNext" delay="0">
                    <p:tgtEl>
                      <p:sldTgt/>
                    </p:tgtEl>
                  </p:cond>
                </p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      {bld_p_xml}
    </p:timing>"""
    return parse_xml(timing_xml.strip())


class CardInfo:
    """
    Represents a card / bounding box container and its text components.
    
    Attributes:
        card_shape: The container background auto-shape (e.g. rounded rectangle).
        tb_shape: The textbox shape sitting inside the card.
        header_paras: Number of leading non-bullet paragraphs (title, subtitle, etc.).
        is_bullet_card: If True and slide has <= 2 cards, bullet paragraphs animate on click.
        extra_shapes: Additional companion shapes to reveal alongside this card.
    """
    def __init__(self, card_shape, tb_shape, header_paras=1, is_bullet_card=True, extra_shapes=None):
        self.card_shape = card_shape
        self.tb_shape = tb_shape
        self.header_paras = header_paras
        self.is_bullet_card = is_bullet_card
        self.extra_shapes = extra_shapes or []


def apply_card_animations(slide, cards):
    """
    Applies animations to a slide:
    - If > 2 bounding boxes (e.g. 3 or 4 cards): Every bounding box appears on click as a unit.
    - If 1 or 2 bounding boxes: Every bounding box container appears on click,
      and each bullet point inside it appears on its own sequential click.
    """
    if not cards:
        return
        
    steps = []
    bld_shapes = set()
    
    if len(cards) > 2:
        for card in cards:
            click_targets = [{'spid': card.card_shape.shape_id}]
            if card.tb_shape:
                click_targets.append({'spid': card.tb_shape.shape_id})
            for extra in card.extra_shapes:
                click_targets.append({'spid': extra.shape_id})
            steps.append(click_targets)
    else:
        for card in cards:
            if card.is_bullet_card and card.tb_shape:
                bld_shapes.add(card.tb_shape.shape_id)
                num_paras = len(card.tb_shape.text_frame.paragraphs)
                
                # Step 1: Card container and header paragraph(s) appear on click
                first_click = [{'spid': card.card_shape.shape_id}]
                if card.header_paras > 0 and num_paras > 0:
                    first_click.append({
                        'spid': card.tb_shape.shape_id,
                        'p_range': (0, min(card.header_paras - 1, num_paras - 1))
                    })
                for extra in card.extra_shapes:
                    first_click.append({'spid': extra.shape_id})
                steps.append(first_click)
                
                # Subsequent clicks: Each bullet paragraph appears on click
                for p_idx in range(card.header_paras, num_paras):
                    steps.append([{
                        'spid': card.tb_shape.shape_id,
                        'p_range': (p_idx, p_idx)
                    }])
            else:
                # Non-bullet card (e.g. code block or diagram placeholder)
                click_targets = [{'spid': card.card_shape.shape_id}]
                if card.tb_shape:
                    click_targets.append({'spid': card.tb_shape.shape_id})
                for extra in card.extra_shapes:
                    click_targets.append({'spid': extra.shape_id})
                steps.append(click_targets)
                
    timing_element = generate_timing_element(steps, bld_shapes=bld_shapes)
    slide.element.append(timing_element)
