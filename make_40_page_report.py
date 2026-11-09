from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path('/Users/kartikeykandpal/Documents/Codex/2026-10-22/mak')
OUT=ROOT/'outputs'/'Pizzeria_Bello_College_Project_Report_40_Pages.docx'
WORK=ROOT/'work'
CREAM='F7F2E9'; INK='29261F'; RED='BD4932'; GREEN='566648'; MUTED='827D70'; PALE='EAE3D6'; WHITE='FFFFFF'

pages=[
('Project overview','Purpose and context',[
'Pizzeria Bello is a responsive website concept for a neighborhood pizza restaurant. The project demonstrates how a small restaurant can present its identity, menu, story, location, and opening hours through a single, easy-to-navigate web page. It is designed as a front-end college project using HTML, CSS, and a small amount of JavaScript.',
'The website uses a warm visual identity built around cream backgrounds, tomato-red accents, muted green details, and editorial typography. Its main visitor journey is straightforward: learn what the restaurant offers, browse a few featured pizzas, and find practical visit information. The page is organized into a hero section, menu, story section, visit details, and footer.',
'This report records the project context, requirements, design choices, implementation structure, quality considerations, and possible next steps. It distinguishes implemented front-end behavior from future features. The restaurant name, address, telephone number, prices, and menu descriptions are sample content and should be replaced before real-world use.'
],['Project type: responsive static restaurant website','Core tools: HTML5, CSS3, JavaScript','Deliverable: one-page site with section navigation','Status: front-end concept; no ordering backend']),
('Introduction','A small business website as a digital front door',[
'A restaurant website often acts as a visitor’s first point of contact. Before deciding where to eat, a customer may want to know what kind of food is served, how much it costs, where the restaurant is located, and whether it is open. A clear website can answer these questions quickly and reduce friction in the decision to visit.',
'Pizzeria Bello is a fictional neighborhood pizzeria used to explore this problem through a practical design and development exercise. The project applies foundational web technologies to a familiar use case. HTML gives the page its structure, CSS presents the visual system and responsive layout, and JavaScript adds a small interaction for the mobile navigation menu.',
'The project is intentionally limited to a static front end. It demonstrates the information and visual experience of a restaurant website without claiming to process orders, reservations, payments, or customer data. This keeps the scope appropriate for a college prototype while leaving a clear path for future development.'
],['A visitor should understand the restaurant within a few seconds.','The menu and visit details should be easy to locate.','The same content should remain usable on narrow screens.']),
('Project submission details','Complete these fields before submission',[
'This page is provided as an editable submission cover sheet. Replace the bracketed prompts with the details required by the college. The report describes a sample website project and does not certify that the work has been accepted, approved, or evaluated by an instructor.',
'Use the project title consistently across the cover, presentation slides, and any source-code submission. If the department requires a specific declaration, certificate, or formatting convention, replace the sample pages in this report with the official format supplied by the college.',
'The author should review the sample restaurant information and either retain it as a clearly fictional concept or substitute approved project content. Confirm that any external photographs, fonts, and icons used in the final website are permitted for the intended academic or public use.'
],['Student name: ______________________________________________','College / university: _________________________________________','Department and course: _______________________________________','Faculty guide: _______________________________________________','Academic year / submission date: ______________________________']),
('Student declaration','Editable declaration template',[
'I, [Student Name], declare that this project report titled “Pizzeria Bello Website Project” has been prepared for academic submission in [Course / Department] at [College / University]. The report describes a front-end website concept developed using HTML, CSS, and JavaScript.',
'I understand that the restaurant identity, menu, prices, address, telephone number, and business details shown in the concept are sample content. They do not represent a verified operating restaurant. Any changes required by my instructor or institution will be made before final submission.',
'I have reviewed the included references and will add any additional sources, acknowledgments, or contribution statements required by the course. This page is a template and should be completed and signed only by the student whose work is being submitted.'
],['Student signature: __________________________________________','Date: __________________________','Place: _________________________'], 'Template page: complete this declaration only if it matches your institution’s requirements.'),
('Project review page','Supervisor or evaluator comments',[
'This page provides a space for feedback on the project. It is intentionally left as a review template and does not claim that a supervisor has reviewed, accepted, or certified the work. The student may retain the blank fields for a formal review meeting or remove this page if the course does not require it.',
'A reviewer may consider whether the project scope is appropriate, whether the design and implementation match the stated objectives, and whether the student has clearly separated implemented features from proposed improvements. The report also identifies project data that is fictional or illustrative.',
'Any grade, signature, institutional seal, or approval statement must be supplied by the authorized instructor or department. Do not present the blank fields as evidence of approval.'
],['Reviewer / faculty guide: ____________________________________','Comments: __________________________________________________','___________________________________________________________','Evaluation / grade: __________________________________________','Signature and date: _________________________________________']),
('Acknowledgement','A short note to personalize',[
'I would like to thank [Faculty Guide Name] for guidance during the planning and development of this college project. Their feedback helped shape the project scope and encouraged a careful approach to layout, content, and usability.',
'I am grateful to [College / Department] for providing the opportunity to apply web development concepts in a practical assignment. I also appreciate the support of classmates, friends, or family members who shared informal feedback during the design process.',
'This acknowledgement is a draft. Before submission, replace the bracketed text with accurate names and contributions. Include only people who actually supported the project, and follow the course instructions on acknowledgements.'
],['Faculty guide: [Name]','Department / institution: [Name]','Additional contributors: [Names, if applicable]']),
('Abstract','Project summary',[
'Pizzeria Bello is a responsive, single-page website concept for a fictional neighborhood pizza restaurant. The project was created to demonstrate how core front-end technologies can organize restaurant information into a clear digital experience. The website introduces the brand, presents selected menu items, shares a short restaurant story, and gives visitors sample location and opening-hour details.',
'The page is built with semantic HTML, a custom CSS design system, and a short JavaScript interaction for expanding and collapsing the navigation on small screens. The visual direction uses a cream, red, and green palette, large editorial headings, food imagery, and responsive layouts. The design includes descriptive page metadata and accessible labels for key controls and image regions.',
'The project is a front-end prototype. It does not connect to a content management system, order platform, payment service, table reservation system, or database. Its restaurant details are fictional sample content. The report documents requirements, design decisions, implementation concepts, and a proposed quality-assurance plan that can be completed before a live demonstration.'
],['Keywords: restaurant website, responsive design, HTML, CSS, JavaScript, user experience']),
('Executive summary','What the project delivers',[
'The project delivers a compact digital storefront for a neighborhood pizzeria. It helps visitors move from discovery to practical information: the hero section introduces the brand, the featured menu gives a quick sense of the food, and the visit section provides sample contact details and hours. The story section adds context and gives the restaurant a more personal voice.',
'The development approach favors simple, readable files over a complex framework. The page structure is contained in index.html, visual rules are in style.css, and mobile navigation behavior is in script.js. The site can be viewed locally by opening the HTML file in a modern browser. External font and photo URLs are used in the design, so those assets depend on internet availability.',
'The next stage should be content verification and hands-on review. A student should confirm all text, check the layout at mobile and desktop widths, test navigation with mouse and keyboard, and replace sample contact details. A genuine order workflow would require additional planning, a trusted service, and appropriate privacy and payment handling.'
],['Primary audience: local diners and first-time visitors','Main task: view menu and plan a visit','Implementation: static client-side site','Future direction: verified content and optional service integrations']),
('Contents','Report map: front matter and chapters',[
'The report is arranged to move from the project context to the implementation and then to review and future development. The first pages introduce the restaurant concept and make room for the information required by an academic submission. Chapters that follow define the problem, objectives, scope, requirements, user journeys, information architecture, and visual design.',
'The implementation chapters describe the roles of HTML, CSS, and JavaScript in the current project. The quality chapters propose ways to review accessibility, responsiveness, content accuracy, and interaction behavior. Because no formal user study or test run is represented in this report, the checklists are framed as work to complete rather than results already achieved.',
'The final pages summarize expected outcomes, limitations, future enhancements, references, a user guide, and a glossary. The report should be adapted to the instructor’s required chapter order, citation style, and submission format.'
],['Front matter: pages 1–7','Project background and planning: pages 8–20','Design and implementation: pages 21–34','Quality, outcomes, and appendices: pages 35–40']),
('Contents continued','Chapter and appendix index',[
'Chapter 1 introduces the need for a restaurant website and describes the fictional Pizzeria Bello concept. Chapter 2 defines the problem, objectives, boundaries, and development method. Chapter 3 describes intended visitors and requirements. Chapter 4 covers information architecture and content. Chapter 5 explains the visual system and responsive design.',
'Chapter 6 documents the static front-end architecture and the HTML, CSS, JavaScript, and asset choices. Chapter 7 covers privacy, security, accessibility, and proposed quality assurance. Chapter 8 reviews the project’s expected contribution, limitations, and possible next steps. The concluding pages include sources, a short visitor guide, and terminology.',
'Suggested figures in the report include the home-page design preview, featured menu preview, navigation structure, interaction flow, and front-end architecture. The two website previews are illustrative reconstructions of the local design, not captures of a live deployed service.'
],['Chapter 1: Introduction and project context','Chapter 2: Problem, objectives, scope, and method','Chapter 3: Users and requirements','Chapter 4: Information architecture and content','Chapter 5: Visual and responsive design','Chapter 6: Front-end implementation','Chapter 7: Quality, privacy, and accessibility','Chapter 8: Outcomes and future work']),
('List of figures and terminology','Quick reference',[
'The report uses a small set of web-development terms. A “responsive layout” adapts its arrangement to the available screen width. A “semantic element” describes the role of content, such as navigation or a main section. A “call to action” is a prominent link that encourages a visitor to take a next step.',
'The site is a static front end: the browser loads files and displays content, but no server-side application handles orders or stores restaurant data. A “breakpoint” is a CSS condition that changes layout rules at a chosen viewport width. The actual design uses a mobile breakpoint to stack columns and show a compact navigation menu.',
'Figure captions identify the site previews and diagrams included in the report. Since the restaurant is fictional and the design is a prototype, figures should be treated as project visuals rather than evidence of a live business.'
],['Figure 1: Pizzeria Bello home-page design preview','Figure 2: Featured pizza menu design preview','Figure 3: Information architecture and visitor journey','Figure 4: Static front-end architecture','HTML: page structure  |  CSS: presentation  |  JS: behavior']),
('Chapter 1 — Introduction','The role of a restaurant website',[
'A restaurant website has a practical job: it should help a visitor decide whether the restaurant fits the occasion and then make the next step obvious. A diner may arrive from a search result, a recommendation, or a social post. The page should quickly communicate cuisine, atmosphere, menu highlights, location, hours, and a way to contact the restaurant.',
'The Pizzeria Bello project treats the website as a simple, brand-led information experience. It is not designed to replace a restaurant’s staff or its point-of-sale system. It provides the basic information that helps a visitor prepare for a visit, while a future service integration could support online ordering or bookings.',
'The project is a learning exercise as well as a design artifact. It connects semantic markup, responsive layout, visual hierarchy, content writing, and a small interactive component. These skills can be applied to many small-business websites beyond the restaurant example.'
],['Introduce the restaurant identity','Show a small set of menu highlights','Make location and contact information easy to scan']),
('Digital presence for local dining','Context and design opportunity',[
'Many visitors use a phone to compare local dining options. A page that depends on wide desktop layouts or hides practical information below decorative content can create unnecessary effort. A small-screen-first perspective helps keep menu labels, location details, and navigation available without overwhelming the screen.',
'Pizzeria Bello uses a one-page structure to keep related information together. The menu is linked from the top navigation and from the hero call to action. The location and hours appear in a dedicated visit section. This approach is suitable for a compact concept site where the amount of content is limited and the primary actions are known.',
'The design also accounts for the character of a restaurant brand. Warm colors, large typography, and food images make the page feel more personal than a plain directory listing. These visual choices support the communication goal, but they should never obscure opening hours, prices, contact links, or accessibility.'
],['Mobile access is an important design constraint.','One-page navigation supports a short information journey.','Visual identity should reinforce, not replace, useful details.']),
('Pizzeria Bello concept','Brand, audience, and content assumptions',[
'Pizzeria Bello is a fictional pizzeria with a friendly neighborhood personality. Its visual voice is informal and welcoming. The copy emphasizes fresh pizza, shared meals, and a relaxed visit. The business name, founding date, menu, location, prices, and telephone number are fictional examples used to make the concept feel complete.',
'The primary audience is a local diner deciding where to eat. Secondary audiences include someone planning a group meal and a first-time visitor looking for directions or hours. These audiences are assumed for the design exercise; they are not the result of a formal survey or interview study.',
'Before a real launch, the restaurant owner would need to approve the brand, photographs, menu descriptions, allergen information, prices, hours, and contact details. The current prototype does not validate any business fact. Clear ownership of the content is necessary to keep customer-facing information accurate.'
],['Brand personality: warm, generous, relaxed','Audience: local diners and first-time visitors','Content status: fictional and for demonstration only']),
('Problem statement','The information gap the site addresses',[
'A potential diner may not know what a restaurant serves, whether it has options they enjoy, or where to find it. If menu highlights and practical details are difficult to locate, visitors may abandon the decision or need to search elsewhere. A small restaurant benefits from presenting these answers in a clear and consistent format.',
'The project problem is therefore framed as an information-design challenge: how can a fictional neighborhood pizzeria communicate its identity, menu, and visit details in one responsive page? The solution should be easy to scan, visually distinctive, and simple to update at the file level.',
'This statement defines the target for the prototype rather than proving a commercial problem through research. A larger project could validate it by interviewing restaurant owners and customers, examining current discovery journeys, and comparing the information visitors need across devices.'
],['Challenge: present useful restaurant information with minimal friction.','Design question: how can a one-page site guide a visitor from discovery to visit planning?']),
('Project objectives','What the prototype aims to demonstrate',[
'The project aims to produce a complete front-end concept that communicates the essential information of a neighborhood pizzeria. It should demonstrate semantic document structure, a consistent visual design, responsive layout behavior, and a small interaction that makes mobile navigation usable.',
'The objectives are deliberately concrete. Each objective can be checked by reviewing the files and observing the page. This report avoids claiming outcomes that would require a formal experiment, such as increased restaurant sales, improved conversion, or measured user satisfaction.',
'A student can use the objectives to structure a presentation: explain the user need, show how the site responds, describe implementation choices, and identify what would be needed to turn the prototype into a real service.'
],['Build a clear restaurant landing page.','Present three featured pizzas with descriptions and sample prices.','Provide a restaurant story, location, hours, and contact link.','Adapt the layout and navigation for smaller screens.','Document limitations and propose a practical next phase.']),
('Project scope','Included work and boundaries',[
'The current scope includes a single HTML page with linked sections, a stylesheet containing desktop and mobile presentation rules, and JavaScript that toggles the mobile navigation. The page contains a hero section, featured menu cards, a story section, visit details, and a footer. The page metadata includes a title and short description.',
'The prototype excludes a backend, user accounts, a shopping cart, order submission, online payment, reservation management, inventory, analytics, and a content management system. It also does not provide a verified allergen database or delivery coverage information. These functions require additional product, technical, privacy, and operational decisions.',
'Keeping these boundaries explicit helps the reader understand what is currently demonstrated. A visible “Order online” style call to action is a design element; it is not evidence that a working order workflow exists. Any public release should adjust that link or connect it to a verified service.'
],['Included: static content, menu cards, anchors, mobile menu toggle','Excluded: transactions, data storage, bookings, live menu management','Assumption: one restaurant location and a small fixed menu']),
('Development method','A small iterative workflow',[
'A suitable workflow for this project begins with content and requirements, then moves through page structure, visual styling, interaction, and review. The site can be built incrementally: create the semantic sections, add the design system, check the layout at several widths, add navigation behavior, and review the content and links.',
'Iteration is useful because design and technical decisions affect each other. For example, a large hero image may work on desktop but consume too much vertical space on a phone. A menu label may appear clear in a desktop header yet be difficult to tap in a compact menu. Review should therefore happen throughout development rather than only at the end.',
'For a formal academic project, the student can keep brief notes about requirements, design changes, and review findings. The current report describes a proposed method and the structure of the supplied files; it does not claim that a documented user study or formal test campaign has been completed.'
],['1. Define the audience and primary tasks.','2. Draft content and page sections.','3. Build HTML structure and CSS presentation.','4. Add and review mobile navigation behavior.','5. Review content, layout, links, and accessibility.']),
('Stakeholders and visitor profiles','Who the site is designed to help',[
'The restaurant owner or project sponsor needs the page to reflect the brand and publish accurate information. The visitor needs quick answers about the food and the visit. The student developer needs an implementation small enough to explain and maintain. An instructor evaluates whether the project demonstrates the required learning outcomes.',
'A first-time visitor may arrive with little context and should understand the restaurant offer quickly. A returning customer may want the menu or hours without rereading the story. A person using a phone may need a compact navigation menu and tappable links. These profiles help the design focus on realistic tasks without implying a formal research sample.',
'If the project continues, feedback should be collected from people who represent these situations. The student can ask participants to find a menu item, locate the address, and explain what the website offers. Observations should be documented consistently and shared only with appropriate consent.'
],['First-time visitor: learn the offer and atmosphere.','Returning visitor: check a menu highlight or opening time.','Mobile visitor: navigate and contact the restaurant using a phone.','Owner or editor: keep business content accurate.']),
('Functional requirements','What a visitor can do',[
'Functional requirements describe observable behavior. The site should load its main content, provide links to the home, menu, story, and visit sections, display featured pizzas with descriptions and prices, and show a restaurant address, hours, and phone link. These are content and navigation behaviors rather than backend operations.',
'On smaller screens, a menu button should open and close the navigation links. The button should expose its expanded state to assistive technology and update its accessible label as the menu changes. Selecting a section link should navigate to its target and close the compact navigation.',
'The requirements reflect the current code structure and can be reviewed in a browser. External order processing is outside the current functionality. If a future order button is introduced, its destination and behavior should be verified before the site is presented as operational.'
],['FR-01: show the restaurant’s introductory hero section.','FR-02: display featured menu items and sample prices.','FR-03: provide section links for menu, story, and visit details.','FR-04: show address, opening hours, and a telephone link.','FR-05: support open and close actions for mobile navigation.']),
('Non-functional requirements','Quality characteristics',[
'Non-functional requirements describe qualities that shape the experience. The layout should adapt to typical phone and desktop widths, text should remain readable, and interactive elements should have meaningful names and visible focus behavior. The design should keep the restaurant’s core information easy to scan.',
'Maintainability is supported by separating HTML, CSS, and JavaScript into named files. This separation makes it easier to find content, presentation rules, and the navigation interaction. The page should avoid unnecessary dependencies so that a student can explain how it works.',
'Performance and reliability depend partly on external font and photo services. If those resources are unavailable, the page should still communicate its text and retain fallback fonts. A future production version should use optimized, licensed assets and check loading behavior under slower connections.'
],['Usability: clear labels and predictable section links.','Compatibility: standard browser features and CSS fallbacks.','Maintainability: separate HTML, CSS, and JavaScript files.','Resilience: meaningful content remains if remote assets fail.']),
('Use cases and visitor journeys','From arrival to next step',[
'A first-time visitor may land on the home section, read the short introduction, and select “Explore the menu.” The page moves to the featured menu section, where the visitor compares three pizzas. If interested, the visitor can continue to the story or visit details.',
'A returning customer may use the top navigation to move directly to the menu or visit section. On a narrow screen, the user opens the mobile menu, selects a section link, and continues reading. The navigation closes after a link is selected so it does not cover the destination content.',
'The visit journey ends with sample directions, opening hours, or a telephone link. In the prototype, these details are placeholders. A deployment should confirm that links point to the correct location and number, and should not imply online ordering unless that service is actually connected.'
],['Journey A: Home → Featured menu → Visit details → Directions or call.','Journey B: Mobile menu → Story → Menu → Continue browsing.','Use case boundary: the website directs visitors; it does not complete an order.']),
('Information architecture','How the one-page site is organized',[
'The site uses a single document with section identifiers. The header contains the restaurant mark and primary navigation. The main content is ordered as hero, menu, story, and visit information. The footer closes the page with the brand and a return-to-top link. This sequence introduces the restaurant before asking visitors to act.',
'The information hierarchy should make the most important task easy to find. Menu and visit details are both prominent because they answer common visitor questions. The story section adds emotional context but does not interrupt access to practical information.',
'One-page navigation is suitable while the amount of content remains compact. If the restaurant adds a full menu, allergen guidance, catering details, or multiple locations, separate pages may become easier to navigate and maintain.'
],['HEADER: brand mark + primary navigation','MAIN: hero → menu → story → visit','FOOTER: brand statement + return-to-top link'],'diagram'),
('Content strategy','What the page says and why',[
'The page content uses short blocks that can be scanned quickly. The hero introduces the restaurant with a short message and a link to the menu. Each menu card combines a name, brief ingredient description, and sample price. The story section describes the restaurant’s intended character, and the visit section contains location, hours, and a phone action.',
'Content should be written for clarity before decoration. Menu descriptions should use accurate ingredient names and avoid claims that cannot be verified. Prices should include the correct currency and be updated when they change. Opening hours should state exceptions where relevant, such as holidays or seasonal changes.',
'Photography and copy should reinforce one another. A photo should depict the food or atmosphere that the text describes. If the image is decorative, it can be treated differently from an image that conveys a specific product or detail. Alternative text should describe meaningful image content without repeating surrounding copy.'
],['Use specific, accurate menu descriptions.','Keep phone, address, and hours current.','Separate business facts from promotional language.']),
('Visual design system','Color, type, and spacing',[
'The visual system uses cream as the main background, dark charcoal for primary text, tomato red for emphasis, and muted green for supporting labels. A restrained palette keeps the page warm while preserving contrast between content and background. The colors are repeated across headings, buttons, small labels, and decorative elements.',
'Large serif headings give the brand a more editorial feel, while a sans-serif face is used for navigation and body text. The layout uses generous margins and spacing to separate sections. The design relies on hierarchy and whitespace rather than dense borders or heavy decorative framing.',
'Remote web fonts are included as progressive enhancement. A fallback stack should be retained so that the text remains legible if the external font service is blocked. Before a public deployment, the developer should check contrast ratios, font loading, and license terms for all selected assets.'
],['Primary background: warm cream','Accent: tomato red','Supporting color: muted green','Typography: serif headlines and sans-serif body text']),
('Home page design','Hero, message, and first action',[
'The hero section is the first major content block. Its headline, “A little slice of Italy,” gives the restaurant an immediate identity. A short supporting paragraph explains the offer in plain language, while the menu call to action gives visitors a direct route to the featured pizzas.',
'The design balances text and food imagery in two columns on wide screens. The circular oven stamp and caption add a small brand detail. The hero is not intended to carry every business detail; it introduces the restaurant and gives the visitor a reason to continue.',
'On narrower screens, the columns stack so the text remains readable and the image remains visible. The small-screen layout should be reviewed to confirm that the call to action remains near the message and that the image does not push all useful content too far below the fold.'
],['Figure 1. Home-page design preview','The visual is an illustrative reconstruction of the page layout.'], 'image1'),
('Featured menu design','Making food options scannable',[
'The featured menu section presents three sample pizzas as consistent cards. Each card repeats the same content pattern: image, category tag, item name, ingredient description, and price. This repetition helps visitors compare choices without learning a new layout for each item.',
'The cards use short descriptions to fit the one-page concept. A production menu may require additional details, such as size options, allergens, dietary markers, or full ingredient lists. These details should be accurate and should not be inferred from a food photograph.',
'The menu section includes a link to a fuller menu, but the current project does not provide an external menu service. A final demonstration should make this clear or connect the link to an actual page. Sample prices in the project are illustrative and must be replaced before publication.'
],['Figure 2. Featured menu design preview','The three items and prices are sample project content.'], 'image2'),
('Story and visit sections','Trust and practical information',[
'The story section adds a human voice to the brand. It describes a simple idea—good food shared with other people—rather than making unsupported claims about history or ingredients. The sample origin story should be replaced with accurate information if the concept represents a real business.',
'The visit section answers practical questions: where the restaurant is, when it is open, and how to call. The address and phone number in the sample are placeholders. In a real website, the directions link should be checked against the correct map pin, and the telephone link should contain the correct country and area code.',
'Contact information should be easy to find and easy to use. The page uses clear labels and places the details together. A future version could provide a contact form, but it would need validation, spam handling, privacy information, and an appropriate destination for submissions.'
],['Story: introduce the restaurant’s voice.','Visit: show location, hours, and call action.','Footer: reinforce the brand and offer a return path.']),
('Responsive layout','Adapting to screen width',[
'The desktop design places the hero content beside the main image, presents menu cards in three columns, and groups visit information into two columns. At smaller widths, the page changes to a single-column flow, menu cards stack vertically, and the visit details become easier to read without horizontal scrolling.',
'Responsive behavior is implemented with CSS media queries. The current breakpoint is an implementation choice rather than a universal rule. It should be reviewed against the actual content and tested at intermediate widths, where navigation, card widths, or headings may wrap unexpectedly.',
'The mobile navigation is hidden until the visitor opens it. Its state should remain clear, and menu links should be large enough to tap comfortably. Images should crop intentionally and preserve the focal point. A future optimization could provide responsive image sizes so the phone does not download a larger image than necessary.'
],['Wide: two-column hero and three menu cards.','Narrow: stacked hero and single-column menu.','Review: narrow phones, larger phones, tablets, and desktop widths.']),
('Accessibility considerations','Designing for more visitors',[
'Accessibility should be part of the design and review process. The site includes a descriptive document title, a viewport declaration, semantic page landmarks, and an accessible label for the mobile menu button. The button exposes its expanded state so that assistive technology can communicate whether the compact menu is open.',
'A complete accessibility review should also examine keyboard access, focus visibility, contrast, heading order, link purpose, image alternatives, and screen-reader behavior. The presence of a label alone does not establish that the whole page conforms to an accessibility standard.',
'The project can use WCAG 2.2 as a reference for organizing an accessibility review. Conformance requires evaluating the full page and relevant responsive variations, not merely selecting a few techniques. The student should describe any review as a check performed, planned, or not yet performed with appropriate accuracy.'
],['Review keyboard navigation and focus visibility.','Provide useful alternative text for meaningful images.','Check contrast and link names.','Reference: W3C Web Content Accessibility Guidelines 2.2.']),
('Front-end architecture','A simple static site',[
'The implementation uses a browser-only architecture. The browser reads the HTML document, loads the stylesheet and JavaScript file, then requests remote image and font assets referenced by the page. There is no application server, database, order API, or payment processor in the current version.',
'This architecture is appropriate for demonstrating a compact informational page. It is easy to open locally and makes the project files understandable to a beginner. The tradeoff is that changing menu content requires editing the source files, and dynamic business functions are not available.',
'If the site becomes a real restaurant service, the architecture should be reconsidered. A content management system could make menu updates easier, while an ordering provider could handle transaction workflows. The student should choose integrations only after confirming business requirements and privacy responsibilities.'
],['Visitor browser → index.html','index.html → style.css + script.js','Browser → external photo and font URLs','No server-side storage or transaction flow'],'diagram'),
('HTML structure','Meaningful page regions',[
'The HTML file defines the document language, character encoding, viewport metadata, page description, and title. The body begins with an announcement strip and header. The header contains a brand link, compact menu control, and primary navigation. Main content is divided into labeled sections using stable identifiers that match the navigation links.',
'Semantic elements such as header, nav, main, section, article, and footer express the role of each region. Headings provide a reading hierarchy. Menu cards use article-like groupings, while the visit details provide contact information. These structural choices help both readers and assistive technologies understand the page organization.',
'The markup should be checked for matching IDs, unique section targets, valid nesting, and descriptive link text. If the content grows, a table of contents or additional pages may make it easier to find information than extending a single document indefinitely.'
],['Document metadata: title, description, viewport, language','Landmarks: header, navigation, main, footer','Sections: home, menu, story, visit']),
('CSS presentation','Reusable rules and responsive styles',[
'The stylesheet defines design tokens as CSS custom properties for colors and font families. These tokens make the palette easier to maintain because a shared color can be adjusted in one place. Global rules apply box sizing and set the base font and background, while component rules style the header, buttons, cards, story panel, visit details, and footer.',
'The desktop layout uses CSS Grid and Flexbox to arrange content. Media queries adjust the number of columns, spacing, and navigation behavior on narrow screens. Images are assigned through background-image rules and cropped with cover behavior. This is visually convenient but may require additional accessible text treatment when the image carries meaning.',
'A future refinement could move informative images into HTML image elements with responsive sources and explicit alternative text. CSS should also be reviewed for focus styles, reduced motion preferences, and readable fallback behavior.'
],['Tokens: palette, body font, display font','Layout: grid and flex patterns','Responsive rule: stack content and compact navigation']),
('JavaScript interaction','The mobile navigation toggle',[
'The JavaScript file selects the navigation panel and menu button, then listens for a click. On each click, it toggles the open class, updates the button’s aria-expanded state, changes its accessible label, and swaps the icon between a menu mark and a close mark.',
'When a visitor selects a navigation link, a second event handler closes the mobile menu and resets the button state. This gives the visitor a predictable interaction on small screens. The rest of the page is static HTML and CSS; JavaScript does not fetch menu data or submit an order.',
'The interaction can be improved through progressive enhancement and keyboard review. The menu should remain understandable if scripts fail, and the button should behave consistently with keyboard activation. If more interactions are added, state changes should be documented and tested for assistive technology.'
],['State: closed → open → closed','Accessibility state: aria-expanded and updated label','Scope: navigation only; no order or payment behavior']),
('Images and external assets','Photography, fonts, and loading',[
'The website references photographs hosted by an external image service and web fonts hosted externally. This keeps the project files small, but it means that the appearance can change if a remote asset is moved, blocked, or unavailable. The page should remain understandable even when images or fonts fail to load.',
'Before publishing, the project owner should confirm the license and permitted use for each photograph and font. Descriptive alternative text should be supplied for meaningful images, while decorative images should not create redundant announcements. If background images remain in use, a text alternative should be available where needed.',
'Performance improvements may include choosing appropriately sized images, compressing files, using modern formats, and loading below-the-fold images lazily. The project should avoid downloading large full-resolution images when a smaller display size would be sufficient.'
],['Confirm licensing and attribution needs.','Add useful alternatives for informative visuals.','Consider local assets and responsive image sizing.']),
('Privacy, security, and responsible content','Prototype boundaries',[
'The current website does not collect or store visitor information. It does not include accounts, forms, payment data, analytics, or order history. This reduces the privacy surface of the prototype, but it does not remove the need to review externally loaded resources or future integrations.',
'The sample contact details and location are fictional placeholders for a project concept. They should not be presented as the details of an operating restaurant. Menu content should be checked for accuracy, especially ingredient and allergen statements. Photographs should not imply a product that the restaurant does not actually serve.',
'If a future version adds forms, tracking, or online checkout, the owner should identify what data is collected, why it is needed, how long it is retained, and who can access it. Payment processing should be delegated to an appropriate provider rather than implemented as a student-built form that stores card details.'
],['Current state: no visitor-data collection or payment flow.','Future data features require privacy review and clear notice.','Verify content, image rights, and business details before launch.']),
('Quality assurance plan','Checks to complete before demonstration',[
'Quality assurance should cover content, links, layout, interactions, and accessibility. The following plan is proposed for a student review. It is not a record of completed tests. The student should run the checks in the browser and document the actual outcome, browser, device or viewport, and any defect found.',
'Content checks should confirm that headings, menu names, descriptions, prices, address, telephone number, and hours are consistent. Link checks should confirm that section anchors, map directions, phone links, and calls to action behave as intended. Layout checks should cover a range of viewport widths and zoom levels.',
'Interaction checks should confirm that the mobile menu opens, closes, exposes its state, and closes after a section link is selected. Accessibility review should include keyboard-only navigation and assistive technology where available. Any remaining issues should be disclosed in the final presentation.'
],['Proposed status: not recorded as executed in this report.','Record actual findings only after the review is performed.']),
('Responsive review checklist','Proposed manual review',[
'A manual responsive review checks whether content remains readable and usable as the viewport changes. Begin at a narrow phone width, then review a larger phone, tablet, and desktop layout. Resize gradually near the breakpoint to identify awkward wrapping or sudden content jumps.',
'At each width, confirm that the navigation remains available, menu cards fit the screen, headings do not overlap images, and buttons remain tappable. Check that there is no unwanted horizontal scrolling. If text becomes too small or the layout requires excessive zoom, adjust the CSS and repeat the review.',
'This checklist should be completed in a real browser using the current project files. The report includes it as a proposed procedure; it does not state that every screen size or browser has already been verified.'
],['□ Narrow phone: navigation and hero action','□ Phone: menu cards and visit information','□ Tablet: transition between stacked and multi-column layouts','□ Desktop: image crop, spacing, and content balance','□ Zoom and text scaling: no loss of essential content']),
('Expected outcomes and learning','What the project demonstrates',[
'The expected outcome is a coherent restaurant landing page that makes its central information easy to find. The project demonstrates how structure, visual hierarchy, responsive layout, and interaction can work together in a small web experience. It also gives the student a concrete example for explaining the relationship between HTML, CSS, and JavaScript.',
'The project encourages practical decisions about content ownership and scope. A polished front end can suggest that a site is ready for business, so the report explicitly identifies what is fictional and which functions are not implemented. This distinction is important when presenting a prototype to an instructor or a potential stakeholder.',
'The educational value comes from building and explaining the artifact, not from claiming commercial performance. Measures such as conversion rate, user satisfaction, and increased sales would require a deployed site, defined measurement methods, and appropriate participant or business data.'
],['Expected learning: semantic markup, layout, responsive behavior, interaction','Expected deliverable: explainable front-end concept and documented next steps']),
('Limitations and future work','What should happen next',[
'The prototype has several limitations. It uses fixed sample content, a single page, remote image and font services, and a static menu. It does not support online orders, table booking, menu editing, multiple languages, dietary filtering, or content updates without editing files. The sample location and contact details must not be used as real business details.',
'A practical next step is to validate the content and visual direction with the restaurant owner or course reviewer. The student can then review the page at multiple widths, check keyboard behavior, optimize imagery, and replace sample data. If the site will be public, hosting, a domain, asset licensing, and accessibility review should be planned.',
'Further features should be prioritized based on actual need. A complete online ordering experience is a separate project with operational, privacy, and payment requirements. It should not be represented as a minor button change.'
],['Near term: verify content, links, and responsive presentation.','Later: consider CMS, full menu, multilingual support, or ordering integration.','Before launch: confirm ownership, hosting, privacy, and accessibility needs.']),
('Conclusion and references','Summary and learning sources',[
'Pizzeria Bello demonstrates a focused way to present a fictional restaurant online. Its one-page structure introduces the brand, highlights the menu, shares a story, and surfaces visit details. The implementation uses HTML, CSS, and JavaScript to create a responsive front-end concept with a mobile navigation interaction.',
'The project should be presented as a prototype rather than a live ordering service. Its details are sample content, its visual previews are illustrative, and its quality-assurance steps remain proposed until the student performs and records them. Clear communication about these boundaries strengthens the academic report.',
'The following sources provide foundational guidance for the web technologies and accessibility concepts discussed in this report. The student may add course materials and any additional sources actually consulted, following the citation style required by the institution.'
],['[1] MDN Web Docs, Web developer guides: https://developer.mozilla.org/en-US/docs/MDN/Guides','[2] MDN Web Docs, HTML guides: https://developer.mozilla.org/en-US/docs/Web/HTML/Guides','[3] W3C, Web Content Accessibility Guidelines (WCAG) 2.2: https://www.w3.org/TR/WCAG22/']),
('Appendix A — Visitor guide','How to use the Pizzeria Bello page',[
'Start at the top of the page to read the short restaurant introduction. Select “Explore the menu” or “Our menu” to jump to the featured pizzas. Each card shows a name, short ingredient description, and sample price. The full-menu link is a concept link and should be connected to a real menu before launch.',
'Use “Our story” to read the restaurant introduction, then select “Find us” to reach the location, hours, and contact details. The map link is intended to open directions, and the phone link is intended to start a call on supported devices. The current address and number are examples only.',
'On a narrow screen, select the menu button to display navigation links. Choose a destination to close the menu and jump to that section. Use the browser’s back-to-top link in the footer to return to the beginning.'
],['For visitors: menu → featured dishes; find us → location and hours.','For the student: replace all sample business information before public use.']),
('Appendix B — Glossary and source files','Key terms and project handoff',[
'HTML (HyperText Markup Language) structures page content. CSS (Cascading Style Sheets) controls visual presentation and layout. JavaScript adds browser-side behavior. A viewport is the visible area of a browser window. A breakpoint is a condition that applies a different style at a particular screen width.',
'A semantic landmark identifies a meaningful page region. Alternative text is a text replacement for an image when that image conveys information. A static site serves files without a custom application server managing data. A content management system provides an interface for editing website content without directly changing source code.',
'The project handoff consists of index.html, style.css, and script.js. Open index.html in a browser to view the static page. To customize the concept, edit the text and links in the HTML, adjust visual rules in the CSS, and update the navigation behavior only if the menu interaction changes.'
],['Files: index.html — content; style.css — presentation; script.js — mobile menu.','Reminder: sample name, menu, prices, address, phone, hours, and links require review.'])
]

# The supplied college checklist determines the final report order and length.
page_map={entry[0]:entry for entry in pages}
page_map.pop('Project overview')
page_map.pop('Introduction')
page_map.pop('Executive summary')
page_map.pop('Contents continued')
page_map.pop('List of figures and terminology')
page_map.pop('Digital presence for local dining')
page_map.pop('Stakeholders and visitor profiles')
page_map.pop('Expected outcomes and learning')
page_map.pop('Appendix A — Visitor guide')
page_map.pop('Appendix B — Glossary and source files')
page_map['Certificate']=('Certificate','Supervisor or evaluator review template',[
'This certificate page is a blank institutional template. It is not a statement that the project has been approved or evaluated. Ask the department whether it requires a specific certificate format, then replace this page with the authorized wording and signature fields.',
'The report describes a responsive front-end website concept for a fictional pizzeria. The reviewer may assess the project against course outcomes such as requirement analysis, system design, implementation, and documentation. Any approval date, grade, seal, or signature must be completed by the authorized person.',
'Do not submit this blank page as evidence of certification. The student should confirm the correct title, student details, supervisor name, and institutional format before printing or sharing the report.'
],['Project title: Pizzeria Bello Website Project','Student: [Name]  |  Course: [Course]','Faculty guide: [Name]  |  Date: [Date]','Authorized signature: _______________________________________'], 'Complete only with your department’s approved certificate text.')
page_map['Feasibility study']=('Feasibility study','Technical, operational, economic, schedule, and legal review',[
'Technical feasibility is high for the current prototype because it uses browser-supported HTML, CSS, and JavaScript without a framework or server. The project files can be opened locally and explained separately. External images and fonts introduce an internet dependency, so local or licensed assets should be considered for a reliable public release.',
'Operational feasibility is suitable for a small informational website. A visitor can browse the menu and locate sample contact details. The static menu is not convenient for frequent updates, and order processing is outside the current system. A restaurant owner would need a simple content update process or a content management system if the menu changes regularly.',
'Economic feasibility cannot be given as a real budget because hosting, domain, asset licensing, and labor rates were not specified. For a classroom prototype, the implementation uses standard front-end files and can be demonstrated without paid backend services. Schedule feasibility is reasonable for a student project when scope remains limited.',
'Legal and ethical feasibility depends on replacing sample data, confirming image and font rights, and accurately describing products and allergens. A future ordering or analytics feature would require separate privacy, security, and payment review.'
],['Assessment: feasible as a static academic prototype.','Condition for public use: verified business details, asset rights, hosting, and accessibility review.'])
page_map['System design overview']=('System design overview','Static information system and key flows',[
'The system is a client-side information site. The visitor loads the page in a browser, reads static sections, and follows links to menu, story, visit details, directions, or a phone action. The browser loads the HTML document, CSS presentation, JavaScript navigation behavior, and referenced image or font assets.',
'There is no authenticated user, application server, database, order submission, or payment service in the current build. The source files are the content store. This design is intentionally small and matches the requirements of a project that presents restaurant information rather than managing restaurant operations.',
'The data-flow diagrams on the following pages describe the current boundary. A conceptual ER model is also included for a possible future content system. That model is a design proposal only; it must not be read as an implemented database.'
],['Primary actor: website visitor','Core processes: render sections, navigate anchors, toggle mobile menu','Data store: static project files','External dependency: remote image and font resources'])
page_map['DFD context diagram']=('DFD context diagram','Level 0 view of the current system',[
'The context diagram treats the website as one process and shows the main exchanges with its environment. A visitor requests the page and interacts with links. The site returns restaurant copy, menu information, location details, and navigation. A browser also requests remote visual assets from an external asset host.',
'No order data, payment details, or personal profile information flows into the project. The directions link opens an external map destination, while the telephone link delegates the call action to the visitor’s device. The project itself does not store the result of either action.',
'The diagram represents the prototype boundary. If ordering, booking, or contact forms are added, new processes, data flows, storage, and security controls must be designed and documented.'
],['Visitor → website: page request and section selections','Website → visitor: restaurant information and navigation','Browser ↔ asset host: image and font requests','Website → map / phone app: outbound visitor action'], 'dfd_context')
page_map['DFD Level 1 diagram']=('DFD Level 1 diagram','Main processes and project-file data store',[
'The level 1 view decomposes the static website into the work visible to a visitor. Process 1 renders the page sections from HTML. Process 2 handles in-page anchor navigation. Process 3 opens or closes the compact navigation on small screens. The source files supply the content and behavior.',
'The browser retrieves external images and fonts when the network and provider allow. These resources are presentation dependencies rather than a restaurant database. The site does not accept or persist an order, reservation, or visitor form submission.',
'This diagram can guide a future extension. A content system would add an editing process and persistent content store. An order system would add a cart, validation, an order service, and a payment provider, each with its own data and security requirements.'
],['1.0 Render page sections  |  2.0 Navigate anchors','3.0 Toggle compact menu  |  D1 Static HTML / CSS / JS','External resources: hosted photography and web fonts'], 'dfd_level1')
page_map['ER diagram']=('ER diagram','Proposed content model for future expansion',[
'A static HTML page does not currently contain a relational database. The entity relationship model below is a proposed way to organize restaurant content if the site is later connected to a content management system. It shows that one restaurant can have many menu items and many opening-hour entries, with one location and contact record in this simplified concept.',
'The model is intentionally small and omits order, customer, payment, and inventory entities because those functions are not in scope. If a real system is designed, the data model should be reviewed with the restaurant and normalized for the selected platform.',
'The examples shown in the diagram are sample values, not verified restaurant records. All business data should be approved before publication.'
],['Current implementation: no database.','Future model only: Restaurant 1—many MenuItem and OpeningHour records.'], 'er')
page_map['Data dictionary']=('Data dictionary','Fields for the proposed content model',[
'This data dictionary explains the key fields represented in the proposed content model. It is a planning aid for a future editable site, not a description of a database that exists in the current front-end project. The examples are fictional and should be replaced with owner-approved values.',
'Field types are conceptual. A production system may use different names or types depending on its content platform. Prices should use a currency-aware numeric representation, and opening times should be stored in a consistent time zone and format.'
],['The data model separates restaurant identity, menu content, location, contact, and opening hours.'], 'data_dictionary')
page_map['Testing plan and test cases']=('Testing plan and test cases','Checks defined for the requested project report',[
'The test plan covers page structure, link targets, responsive metadata, and the mobile navigation behavior. Structural checks and a small mocked interaction harness were run against the supplied project files. The results are summarized on the next page. These checks do not replace a visual review in a real browser.',
'The test cases below are repeatable. A student should also open the page in a browser and review layout at phone, tablet, and desktop widths, use the keyboard to reach links and buttons, and confirm that remote images load. Those browser checks are not represented as completed here.'
],[], 'test_cases')
page_map['Test results']=('Test results','Executed checks and remaining review items',[
'Eight automated checks passed in a local Node.js harness. Four checks inspected the HTML for unique section identifiers, valid in-page link targets, viewport metadata, and a document title. Four checks exercised the mobile navigation handler with a lightweight DOM mock: opening the menu, announcing the expanded state, closing after link selection, and restoring the closed-state label and icon.',
'These results apply only to the checks listed below. They do not establish full browser compatibility, visual correctness, accessibility conformance, or a working ordering service. Manual review in a browser remains necessary before a final classroom demonstration.'
],['Result: 8 checks passed, 0 failed.','Pending: visual review in browser at multiple widths; keyboard and assistive-technology review; link destination confirmation.'], 'test_results')

def with_title(entry,title,subtitle=None):
    return (title, subtitle or entry[1], *entry[2:])

page_map['Certificate']=with_title(page_map['Certificate'],'Certificate')
page_map['Acknowledgement']=with_title(page_map['Acknowledgement'],'Acknowledgment')
page_map['Contents']=('Contents','Required report sections and page numbers',[
'The page numbers below correspond to this 40-page report. Replace this contents page if your college requires an automatic table of contents.'
],[], 'toc')
page_map['Home page design']=('Implementation screenshots — home page','Home section visual preview',[
'The image below is an illustrative preview of the page composition, not a direct capture from a running browser. If your instructor expects execution evidence, replace it with a screenshot of the local site open in a browser before submission.',
'The home section introduces the brand and directs visitors to the menu. The layout balances a short value statement with a large food visual and a clear next step.'
],[],'image1')
page_map['Featured menu design']=('Implementation screenshots — menu','Featured menu visual preview',[
'The image below is an illustrative preview of the menu layout, not a direct capture from a running browser. Replace it with a browser screenshot if your instructor requires evidence from the running implementation.',
'The menu cards show a repeatable pattern for a dish name, short description, sample price, and food image. All menu information is fictional sample content.'
],[],'image2')
page_map['Problem statement']=with_title(page_map['Problem statement'],'Problem definition')
page_map['Problem statement']=with_title(page_map['Problem statement'],'System analysis — problem definition')
page_map['Feasibility study']=with_title(page_map['Feasibility study'],'System analysis — feasibility study')
page_map['Testing plan and test cases']=with_title(page_map['Testing plan and test cases'],'Testing — test cases')
page_map['Test results']=with_title(page_map['Test results'],'Testing — results')
page_map['Limitations and future work']=('Conclusion and future scope','Limitations, recommendations, and next phase',[
'Pizzeria Bello demonstrates how a compact restaurant website can introduce a brand, present menu highlights, and surface visit details through a responsive front end. The project brings together semantic HTML, CSS layout, content hierarchy, and a small JavaScript navigation interaction.',
'The main limitations are the fixed sample content, single-page structure, and dependence on external photo and font resources. The project has no database, order service, booking flow, payment processing, or menu editor. The address, telephone number, prices, menu descriptions, and business story are fictional examples.',
'A practical next phase is to replace sample data with owner-approved content, capture screenshots from the running site, review the page in a browser at multiple sizes, and record keyboard and accessibility findings. Further work could add a full menu, content management, multilingual copy, or an ordering integration after requirements and privacy responsibilities are defined.'
],['Near term: verify content, links, and responsive presentation.','Future options: CMS, full menu, multilingual support, or ordering integration.','Before launch: confirm ownership, hosting, asset rights, privacy, and accessibility.'])
page_map['Conclusion and references']=('References','Sources and further reading',[
'The following primary documentation supports the general web-development and accessibility topics discussed in this report. Add any course notes, tools, or additional material that you personally consulted, using the citation style required by your college.',
'The references are learning resources, not evidence that this prototype conforms to a specific standard. Any accessibility or browser compatibility claim should be based on a review that is actually performed and documented.'
],['[1] MDN Web Docs, Web developer guides: https://developer.mozilla.org/en-US/docs/MDN/Guides','[2] MDN Web Docs, HTML guides: https://developer.mozilla.org/en-US/docs/Web/HTML/Guides','[3] W3C, Web Content Accessibility Guidelines (WCAG) 2.2: https://www.w3.org/TR/WCAG22/'])

ordered_titles=[
'Project submission details','Student declaration','Certificate','Acknowledgment','Abstract','Contents',
'Chapter 1 — Introduction','Pizzeria Bello concept','System analysis — problem definition','Project objectives','System analysis — feasibility study','Project scope','Development method',
'Functional requirements','Non-functional requirements','Use cases and visitor journeys','System design overview','DFD context diagram','DFD Level 1 diagram','ER diagram','Data dictionary',
'Information architecture','Content strategy','Visual design system','Home page design','Featured menu design','Story and visit sections','Responsive layout','Accessibility considerations',
'Front-end architecture','HTML structure','CSS presentation','JavaScript interaction','Images and external assets','Privacy, security, and responsible content','Testing — test cases','Testing — results','Conclusion and future scope','References'
]
ordered_titles[ordered_titles.index('Home page design')]='Implementation screenshots — home page'
ordered_titles[ordered_titles.index('Featured menu design')]='Implementation screenshots — menu'
entries_by_title={entry[0]:entry for entry in page_map.values()}
pages=[entries_by_title[t] for t in ordered_titles]

def shade(cell,color):
    pr=cell._tc.get_or_add_tcPr(); el=OxmlElement('w:shd'); el.set(qn('w:fill'),color); pr.append(el)

def add_page_num(p):
    r=p.add_run('  |  PAGE '); r.font.size=Pt(8); r.font.color.rgb=RGBColor.from_string(MUTED)
    fld=OxmlElement('w:fldSimple'); fld.set(qn('w:instr'),'PAGE'); r._r.addnext(fld)

def add_diagram(doc, variant):
    tbl=doc.add_table(rows=1, cols=1); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.autofit=False; tbl.columns[0].width=Inches(6.2)
    c=tbl.cell(0,0); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; shade(c,PALE)
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    if variant=='diagram':
        for label in ['VISITOR ARRIVES','HOME / HERO','MENU  →  STORY  →  VISIT','CALL OR DIRECTIONS']:
            r=p.add_run(label+'\n'); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=RGBColor.from_string(GREEN if 'VISIT' not in label else RED)
            r=p.add_run('↓\n'); r.font.size=Pt(13); r.font.color.rgb=RGBColor.from_string(MUTED)
    else:
        for label in ['BROWSER','index.html','style.css  +  script.js','REMOTE IMAGES / FONTS']:
            r=p.add_run(label+'\n'); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=RGBColor.from_string(GREEN if 'REMOTE' not in label else RED)
            r=p.add_run('↓\n'); r.font.size=Pt(13); r.font.color.rgb=RGBColor.from_string(MUTED)
    return tbl

def add_grid(doc, headers, rows, widths=None, font_size=8.5):
    tbl=doc.add_table(rows=1,cols=len(headers)); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.style='Table Grid'
    for i,h in enumerate(headers):
        c=tbl.rows[0].cells[i]; shade(c,GREEN); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(1); r=p.add_run(h); r.bold=True; r.font.size=Pt(font_size); r.font.color.rgb=RGBColor.from_string(WHITE)
    for row in rows:
        cells=tbl.add_row().cells
        for i,value in enumerate(row):
            cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p=cells[i].paragraphs[0]; p.paragraph_format.space_after=Pt(1); r=p.add_run(str(value)); r.font.size=Pt(font_size); r.font.color.rgb=RGBColor.from_string(INK)
            if len(tbl.rows)%2==0: shade(cells[i],'F7F2E9')
    if widths:
        for row in tbl.rows:
            for i,w in enumerate(widths): row.cells[i].width=Inches(w)
    return tbl

def add_flow_diagram(doc, kind):
    if kind=='dfd_context':
        add_grid(doc,['EXTERNAL ENTITY','DATA FLOW','SYSTEM PROCESS'],[
            ['Visitor','Page request, link selection','Pizzeria Bello website'],
            ['Pizzeria Bello website','Text, menu, hours, navigation','Visitor'],
            ['Browser','Request / response for fonts and images','External asset host'],
            ['Website','Directions / telephone link','Map or phone application']
        ],font_size=9)
        p=doc.add_paragraph('Boundary note: no order, payment, booking, or personal data is stored by this static prototype.'); p.paragraph_format.space_before=Pt(8)
        for r in p.runs: r.italic=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor.from_string(GREEN)
    elif kind=='dfd_level1':
        add_grid(doc,['PROCESS','INPUT','OUTPUT'],[
            ['1.0 Render page','HTML source and visitor request','Visible restaurant sections'],
            ['2.0 Navigate sections','Anchor-link selection','Menu, story, or visit section'],
            ['3.0 Toggle mobile menu','Button activation','Open / closed navigation state'],
            ['D1 Static files','index.html, style.css, script.js','Content, layout, interaction'],
            ['External asset host','Font / image request','Remote visual assets']
        ],font_size=9)
    elif kind=='er':
        add_grid(doc,['PARENT ENTITY','RELATIONSHIP','RELATED ENTITY'],[
            ['Restaurant (1)','has many  →','MenuItem (many)'],
            ['Restaurant (1)','has many  →','OpeningHour (many)'],
            ['Restaurant (1)','has one  →','Location (one)'],
            ['Restaurant (1)','has one  →','Contact (one)']
        ],font_size=8.5)
    elif kind=='data_dictionary':
        add_grid(doc,['ENTITY.FIELD','TYPE','MEANING','SAMPLE VALUE'],[
            ['Restaurant.name','Text','Displayed restaurant name','Pizzeria Bello'],
            ['Restaurant.description','Text','Short brand introduction','A neighborhood pizzeria'],
            ['MenuItem.name','Text','Menu item title','Margherita'],
            ['MenuItem.price','Decimal + currency','Displayed sample price','14.00 USD (example)'],
            ['OpeningHour.day','Enum / text','Day or day range','Mon–Thu'],
            ['OpeningHour.open / close','Time','Local opening interval','12:00 / 21:00'],
            ['Location.address','Text','Visitor-facing location','Sample address only'],
            ['Contact.phone','Telephone string','Click-to-call destination','Sample number only']
        ],font_size=8.5)
    elif kind=='test_cases':
        add_grid(doc,['ID','TEST CASE','EXPECTED RESULT','CHECK METHOD'],[
            ['TC-01','Section IDs are unique','No repeated IDs','Static HTML check'],
            ['TC-02','In-page links target existing IDs','All targets exist','Static HTML check'],
            ['TC-03','Viewport metadata is present','Mobile viewport declared','Static HTML check'],
            ['TC-04','Document title is present','Non-empty title tag','Static HTML check'],
            ['TC-05','Open mobile menu','Open state and expanded=true','JS interaction harness'],
            ['TC-06','Open state label/icon','Close label and × icon','JS interaction harness'],
            ['TC-07','Select menu link','Menu closes and state=false','JS interaction harness'],
            ['TC-08','Closed state label/icon','Open label and ☰ icon','JS interaction harness']
        ],font_size=8)
    elif kind=='test_results':
        add_grid(doc,['ID','RESULT','OBSERVATION'],[
            ['TC-01','PASS','Section identifiers were unique.'],['TC-02','PASS','All in-page targets were found.'],
            ['TC-03','PASS','Viewport metadata exists.'],['TC-04','PASS','Document title exists.'],
            ['TC-05','PASS','Menu opens and exposes expanded state.'],['TC-06','PASS','Open label and icon are updated.'],
            ['TC-07','PASS','Section-link selection closes the menu.'],['TC-08','PASS','Closed label and icon are restored.']
        ],font_size=8.5)

def add_contents(doc):
    toc_rows=[]
    for i in range(0,len(ordered_titles),2):
        left=f'{i+2:02d}  {ordered_titles[i]}'
        right=f'{i+3:02d}  {ordered_titles[i+1]}' if i+1<len(ordered_titles) else ''
        toc_rows.append([left,right])
    tbl=doc.add_table(rows=0,cols=2); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.autofit=False
    tbl.columns[0].width=Inches(3.25); tbl.columns[1].width=Inches(3.25)
    for row in toc_rows:
        cells=tbl.add_row().cells
        for j,val in enumerate(row):
            cells[j].width=Inches(3.25); p=cells[j].paragraphs[0]; p.paragraph_format.space_after=Pt(2)
            r=p.add_run(val); r.font.size=Pt(8.5); r.font.color.rgb=RGBColor.from_string(INK)

doc=Document(); sec=doc.sections[0]
sec.page_width=Inches(8.27); sec.page_height=Inches(11.7); sec.top_margin=Inches(.68); sec.bottom_margin=Inches(.67); sec.left_margin=Inches(.82); sec.right_margin=Inches(.82)
styles=doc.styles
normal=styles['Normal']; normal.font.name='Arial'; normal.font.size=Pt(10); normal.font.color.rgb=RGBColor.from_string(INK); normal.paragraph_format.space_after=Pt(7); normal.paragraph_format.line_spacing=1.1
for sn,size,color in [('Title',30,INK),('Heading 1',23,GREEN),('Heading 2',14,RED)]:
    st=styles[sn]; st.font.name='Georgia'; st.font.size=Pt(size); st.font.bold=True; st.font.color.rgb=RGBColor.from_string(color)
    if st._element.pPr is not None:
        for n in list(st._element.pPr):
            if n.tag==qn('w:pBdr'): st._element.pPr.remove(n)
header=sec.header.paragraphs[0]; header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
r=header.add_run('PIZZERIA BELLO  •  COLLEGE PROJECT REPORT'); r.font.name='Arial'; r.font.size=Pt(8); r.font.bold=True; r.font.color.rgb=RGBColor.from_string(GREEN)
footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=footer.add_run('PIZZERIA BELLO  •  PROJECT REPORT'); r.font.size=Pt(8); r.font.color.rgb=RGBColor.from_string(MUTED); add_page_num(footer)

# Page 1: cover
p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(64); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('COLLEGE PROJECT REPORT'); r.font.size=Pt(11); r.font.bold=True; r.font.color.rgb=RGBColor.from_string(RED)
p=doc.add_paragraph(style='Title'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(24); p.paragraph_format.space_after=Pt(10); p.add_run('Pizzeria Bello\nWebsite Project')
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(24); r=p.add_run('A responsive restaurant website concept'); r.italic=True; r.font.size=Pt(14); r.font.color.rgb=RGBColor.from_string(MUTED)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run().add_picture(str(WORK/'website-top.png'),width=Inches(5.6))
for line in ['Submitted by: [Student Name]','Course / Department: [Course and Department]','College / University: [Institution Name]','Faculty Guide: [Name]','Academic Year: [Year]']:
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4); r=p.add_run(line); r.font.size=Pt(10)
doc.add_page_break()

# Pages 2–40: all entries are explicit report pages.
for index,(title,subtitle,paras,items,*extra) in enumerate(pages,start=2):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(3)
    r=p.add_run(f'PAGE {index:02d}  /  PROJECT REPORT'); r.font.size=Pt(8); r.font.bold=True; r.font.color.rgb=RGBColor.from_string(RED)
    h=doc.add_paragraph(title,style='Heading 1'); h.paragraph_format.space_after=Pt(2)
    p=doc.add_paragraph(subtitle,style='Heading 2'); p.paragraph_format.space_after=Pt(9)
    for para in paras:
        p=doc.add_paragraph(para); p.paragraph_format.keep_together=True
    if extra and extra[0]=='toc':
        add_contents(doc)
    elif extra and extra[0] in ('dfd_context','dfd_level1','er','data_dictionary','test_cases','test_results'):
        add_flow_diagram(doc,extra[0])
    elif extra and extra[0]=='image1':
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2); p.add_run().add_picture(str(WORK/'website-top.png'),width=Inches(6.0))
        p=doc.add_paragraph('Figure 1. Home-page design preview'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4)
        for r in p.runs: r.italic=True; r.font.size=Pt(8); r.font.color.rgb=RGBColor.from_string(MUTED)
    elif extra and extra[0]=='image2':
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2); p.add_run().add_picture(str(WORK/'website-menu.png'),width=Inches(6.0))
        p=doc.add_paragraph('Figure 2. Featured menu design preview'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4)
        for r in p.runs: r.italic=True; r.font.size=Pt(8); r.font.color.rgb=RGBColor.from_string(MUTED)
    elif extra and extra[0]=='diagram': add_diagram(doc,'diagram')
    elif extra and extra[0]=='architecture': add_diagram(doc,'architecture')
    if items:
        for item in items:
            p=doc.add_paragraph(style='List Bullet'); p.paragraph_format.left_indent=Inches(.25); p.paragraph_format.first_line_indent=Inches(-.12); p.paragraph_format.space_after=Pt(4)
            p.add_run(item)
    handled_extra=('toc','dfd_context','dfd_level1','er','data_dictionary','test_cases','test_results','image1','image2','diagram','architecture')
    if extra and isinstance(extra[0],str) and extra[0] not in handled_extra:
        tbl=doc.add_table(rows=1,cols=1); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; c=tbl.cell(0,0); shade(c,PALE)
        p=c.paragraphs[0]; r=p.add_run(extra[0]); r.italic=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor.from_string(GREEN)
    if index < 40: doc.add_page_break()

doc.core_properties.title='Pizzeria Bello College Project Report'; doc.core_properties.subject='40-page report for a responsive restaurant website concept'; doc.core_properties.keywords='college project, pizza website, HTML, CSS, JavaScript'
doc.save(OUT)
print(OUT)
