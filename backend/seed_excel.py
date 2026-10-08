"""
seed_excel.py
─────────────────────────────────────────────────────────────────────────────
Run this ONCE to pre-populate the Excel database with 100 Q&A for the
Test Manager role (and empty sheets for all other roles).

Usage:
    cd interview-assistant/backend
    python seed_excel.py
─────────────────────────────────────────────────────────────────────────────
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import uuid
from datetime import datetime, timezone
from services import excel_service

# ── 100 Test Manager Q&A ──────────────────────────────────────────────────
TEST_MANAGER_QA = [
    (
        "What is the primary responsibility of a Test Manager?",
        ["Plan, execute, and monitor all testing activities",
         "Define test scope, objectives, and strategy",
         "Organise and manage the test team and resources",
         "Ensure quality standards are maintained throughout",
         "Report testing progress to stakeholders"],
        "A Test Manager is responsible for planning the overall testing strategy, organising test resources, defining test scope and objectives, executing test plans, monitoring test progress against timelines and budgets, and ensuring quality standards are maintained throughout the testing lifecycle."
    ),
    (
        "Define test plan.",
        ["Document outlining testing scope, approach, resources, and schedule",
         "Serves as a roadmap for the entire testing effort",
         "Includes test objectives, entry and exit criteria",
         "Defines roles, responsibilities, and resource requirements",
         "Describes test environment and tools"],
        "A test plan is a comprehensive document that outlines what will be tested, how it will be tested, when testing will occur, who will perform the testing, what resources are required, and the expected outcomes. It serves as a roadmap for the entire testing effort."
    ),
    (
        "What is a test case?",
        ["Step-by-step procedure to validate software functionality",
         "Contains preconditions, inputs, steps, and expected results",
         "Includes pass/fail criteria",
         "Assigned a unique test case ID",
         "Linked to one or more requirements"],
        "A test case is a set of preconditions, inputs, actions, and expected results designed to verify that a specific aspect of the software functions correctly. It includes test case ID, description, steps, expected results, and pass/fail criteria."
    ),
    (
        "Explain the difference between QA and QC.",
        ["QA is process-oriented; QC is product-oriented",
         "QA focuses on prevention of defects",
         "QC focuses on detection of defects through testing",
         "QA involves process audits and standards",
         "QC involves test execution and inspection"],
        "Quality Assurance (QA) focuses on preventing defects through process improvements and planning, while Quality Control (QC) focuses on detecting defects through testing and inspection activities. QA is process-oriented; QC is product-oriented."
    ),
    (
        "What are the main levels of software testing?",
        ["Unit testing — individual components",
         "Integration testing — component interactions",
         "System testing — complete system validation",
         "User Acceptance Testing (UAT) — business requirements",
         "Regression testing — after changes"],
        "Unit testing verifies individual components; Integration testing checks component interactions; System testing validates the complete system; User Acceptance Testing (UAT) confirms the system meets business requirements."
    ),
    (
        "Define regression testing.",
        ["Re-testing after code changes to ensure no new defects",
         "Critical after bug fixes or feature enhancements",
         "Uses previously executed test cases",
         "Can be automated for efficiency",
         "Protects existing functionality"],
        "Regression testing involves executing previously executed test cases after code modifications to ensure that the changes haven't introduced new defects or broken existing functionality. It's critical after bug fixes or feature enhancements."
    ),
    (
        "What is a test metric?",
        ["Quantitative measure of testing progress and quality",
         "Includes defect density, test coverage, pass/fail rate",
         "Supports data-driven decision-making",
         "Tracks team performance and efficiency",
         "Reported to stakeholders regularly"],
        "Test metrics are data points used to measure testing effectiveness, such as test case execution rate, defect density, test coverage percentage, defect detection rate, and overall product quality indicators."
    ),
    (
        "Explain smoke testing.",
        ["Quick test to verify basic system functionality after a build",
         "Sanity check before detailed testing begins",
         "Confirms critical paths work end-to-end",
         "Performed on each new build",
         "Fast and shallow — not exhaustive"],
        "Smoke testing is a preliminary testing type performed after a build to verify that the critical functionalities of the system are working as expected. It's a quick sanity check before detailed testing begins."
    ),
    (
        "What is exploratory testing?",
        ["Ad-hoc testing without predefined test cases",
         "Tester learns and adapts in real time",
         "Relies on domain knowledge and creativity",
         "Effective for finding unexpected issues",
         "Documented as session-based notes"],
        "Exploratory testing is a testing approach where testers actively explore the application without formal test cases, learning as they test and adapting their approach based on findings. It requires domain knowledge and creativity."
    ),
    (
        "Define test coverage.",
        ["Percentage of code or functionality tested",
         "Can be measured as code, requirements, or feature coverage",
         "Expressed as a percentage",
         "Guides where more testing is needed",
         "Higher coverage reduces risk of missed defects"],
        "Test coverage represents the extent to which the application code, features, or requirements have been tested. It can be measured as code coverage, requirements coverage, or feature coverage, expressed as a percentage."
    ),
    (
        "What is the difference between verification and validation?",
        ["Verification: are we building it right?",
         "Validation: are we building the right product?",
         "Verification uses reviews, inspections, static analysis",
         "Validation uses testing, demos, UAT",
         "Both are essential for a quality product"],
        "Verification ensures the software is built correctly according to specifications (reviews, inspections, static analysis), while Validation confirms the software meets user needs and requirements (testing, demos, UAT)."
    ),
    (
        "Explain test prioritization.",
        ["Ranking test cases based on business impact and risk",
         "High-priority tests executed first",
         "Considers frequency of use and criticality",
         "Maximises risk mitigation with limited time",
         "Updated based on changing requirements"],
        "Test prioritization involves ordering test cases based on factors like business impact, risk level, frequency of use, criticality of functionality, and dependencies. High-priority tests are executed first to maximise risk mitigation."
    ),
    (
        "What is defect severity?",
        ["Impact level of a defect on system functionality",
         "Categorised as Critical, High, Medium, Low, Trivial",
         "Set by the tester based on functionality impact",
         "Determines urgency of fix",
         "Different from defect priority"],
        "Defect severity indicates how severely a bug impacts system functionality, typically categorised as Critical, High, Medium, Low, or Trivial based on its effect on functionality and user experience."
    ),
    (
        "Define boundary value analysis.",
        ["Testing at the edges of input ranges",
         "Defects often occur at boundaries",
         "Tests minimum, maximum, and just outside boundary values",
         "Black-box testing technique",
         "Reduces number of test cases needed"],
        "Boundary value analysis is a black-box testing technique that focuses on testing at the edges of input ranges, such as minimum, maximum, and just above/below boundary values, as defects often occur at boundaries."
    ),
    (
        "What is equivalence partitioning?",
        ["Grouping inputs into classes with similar behaviour",
         "Tests one value from each partition",
         "Reduces total number of test cases",
         "Ensures representative coverage",
         "Can be applied to both valid and invalid inputs"],
        "Equivalence partitioning divides input data into groups (partitions) that are expected to behave similarly, testing one value from each partition to reduce test cases while maintaining coverage effectiveness."
    ),
    (
        "Explain the V-model in testing.",
        ["Testing at each SDLC phase matched with development",
         "Each development phase has a corresponding test phase",
         "Forms a V-shape visually",
         "Emphasises early test planning",
         "Defects found earlier are cheaper to fix"],
        "The V-model represents the relationship between development and testing phases, where each development phase has a corresponding testing phase (Unit→Unit Test, Integration→Integration Test, etc.), forming a V-shape."
    ),
    (
        "What is a test harness?",
        ["Tools and code that support test execution",
         "Includes drivers and stubs",
         "Simulates missing components",
         "Provides controlled test environment",
         "Enables isolated unit and integration testing"],
        "A test harness is a collection of software and test data configured to test a program unit by simulating other components, including drivers, stubs, and testing utilities to facilitate testing."
    ),
    (
        "Define path coverage.",
        ["Testing all possible execution paths in code",
         "Most comprehensive form of structural coverage",
         "Accounts for all conditional branches and loops",
         "Can be impractical for large codebases",
         "Ensures no untested paths exist"],
        "Path coverage ensures that every possible execution path through the code is tested at least once, accounting for all conditional branches and loops to achieve comprehensive coverage."
    ),
    (
        "What is statement coverage?",
        ["Testing to execute every code statement at least once",
         "Also called line coverage",
         "Measures percentage of executed statements",
         "Basic but widely used metric",
         "Does not guarantee all branches are tested"],
        "Statement coverage (line coverage) measures the percentage of code statements that have been executed during testing, aiming to execute every line at least once."
    ),
    (
        "Explain branch coverage.",
        ["Tests all decision outcomes (true and false branches)",
         "More thorough than statement coverage",
         "Ensures both outcomes of conditionals are tested",
         "Measured as percentage of branches executed",
         "Helps find logic errors in conditions"],
        "Branch coverage ensures that every branch (true and false outcomes) of conditional statements is executed at least once, providing more thorough coverage than statement coverage."
    ),
    (
        "What is Monte Carlo testing?",
        ["Random testing with statistical analysis",
         "Inputs are randomly selected",
         "Results analysed statistically",
         "Estimates system reliability",
         "No formal test case design required"],
        "Monte Carlo testing involves randomly selecting test inputs and analysing results statistically to estimate system behaviour and reliability without formal test case design."
    ),
    (
        "Define stress testing.",
        ["Testing system performance under extreme load",
         "Determines breaking points",
         "Evaluates recovery capabilities",
         "Tests beyond normal operational capacity",
         "Identifies performance bottlenecks"],
        "Stress testing evaluates system behaviour under extreme, abnormal conditions (high load, limited resources, peak traffic) to determine breaking points and recovery capabilities."
    ),
    (
        "What is load testing?",
        ["Testing system performance under expected user load",
         "Monitors response times and throughput",
         "Tests from normal to peak conditions",
         "Identifies performance degradation",
         "Validates SLA compliance"],
        "Load testing measures how the system performs when subjected to expected user loads, monitoring response times, throughput, and resource utilisation under normal to peak conditions."
    ),
    (
        "Explain performance testing.",
        ["Evaluates speed, scalability, and stability under load",
         "Encompasses load, stress, and endurance testing",
         "Measures responsiveness and resource consumption",
         "Identifies scalability limits",
         "Critical for high-traffic applications"],
        "Performance testing encompasses load, stress, and endurance testing to measure system responsiveness, stability, resource consumption, and scalability under various conditions."
    ),
    (
        "What is usability testing?",
        ["Evaluating user-friendliness and user experience",
         "Focuses on UI design and navigation",
         "Involves real users performing tasks",
         "Identifies UX pain points",
         "Improves overall user satisfaction"],
        "Usability testing assesses how easily users can interact with the application, focusing on user interface design, navigation, user experience, accessibility, and overall satisfaction."
    ),
    (
        "Define compatibility testing.",
        ["Testing application across environments, browsers, OS",
         "Verifies correct function on different hardware",
         "Covers browsers, devices, and network conditions",
         "Includes third-party software compatibility",
         "Critical for web and mobile applications"],
        "Compatibility testing verifies that the application functions correctly across different hardware platforms, operating systems, browsers, devices, network conditions, and third-party software."
    ),
    (
        "What is security testing?",
        ["Testing against security vulnerabilities and threats",
         "Covers SQL injection, XSS, authentication, authorisation",
         "Tests data encryption and access controls",
         "Identifies security weaknesses",
         "Essential for compliance and risk management"],
        "Security testing identifies vulnerabilities, weaknesses, and threats in the system, including testing for SQL injection, XSS attacks, authentication, authorisation, data encryption, and other security concerns."
    ),
    (
        "Explain localization testing.",
        ["Testing application for different languages and regions",
         "Verifies correct translation and cultural adaptation",
         "Checks currencies, date formats, and locale settings",
         "Ensures local regulatory compliance",
         "Includes UI layout adjustments for different scripts"],
        "Localisation testing verifies that the application is properly adapted for different languages, cultures, currencies, date formats, and regional requirements, ensuring local relevance."
    ),
    (
        "What is globalization testing?",
        ["Testing application for international functionality",
         "Supports multiple languages and currencies",
         "Validates date/time and cultural conventions",
         "Ensures no functionality is broken by internationalisation",
         "Foundation for localisation"],
        "Globalisation testing ensures the application can support multiple languages, currencies, date/time formats, and cultural conventions without breaking functionality or user experience."
    ),
    (
        "Define ad-hoc testing.",
        ["Unplanned, informal testing without test cases",
         "Based on tester knowledge and intuition",
         "Useful for finding unexpected issues",
         "No formal documentation required",
         "Complements structured testing approaches"],
        "Ad-hoc testing is informal testing performed without predefined test cases or documentation, where testers use their knowledge and intuition to find defects. It's useful for finding unexpected issues."
    ),
    (
        "What is the role of a Test Manager in risk management?",
        ["Identify testing risks across schedule, resources, and technical areas",
         "Assess probability and impact of each risk",
         "Develop mitigation strategies",
         "Monitor risks throughout the project",
         "Escalate high-impact risks to stakeholders"],
        "A Test Manager must identify risks affecting testing (schedule risks, resource risks, technical risks), assess their probability and impact, and develop mitigation strategies to ensure quality and timeline objectives are met."
    ),
    (
        "How do you manage test team dynamics and conflicts?",
        ["Foster open communication and transparency",
         "Establish clear roles and responsibilities",
         "Address conflicts promptly and fairly",
         "Build team cohesion through training",
         "Provide regular feedback and recognition"],
        "A Test Manager should promote open communication, clarify responsibilities, address conflicts promptly, build team cohesion through training and team building, provide feedback, and ensure a positive work environment."
    ),
    (
        "Explain the concept of test independence.",
        ["Testers separate from developers for objective testing",
         "Reduces bias in defect reporting",
         "Improves defect detection rate",
         "Provides credible quality assessment",
         "Varies from no independence to fully independent teams"],
        "Test independence suggests that testers should be organisationally independent from developers to maintain objectivity, reduce bias, improve defect detection, and provide credible testing results."
    ),
    (
        "What is a traceability matrix?",
        ["Maps requirements to test cases and results",
         "Ensures all requirements are covered by tests",
         "Enables impact analysis for changes",
         "Provides traceability from requirement to defect",
         "Used for audit and compliance evidence"],
        "A requirements traceability matrix (RTM) links each requirement to corresponding test cases, test results, and defects, ensuring all requirements are tested and enabling impact analysis."
    ),
    (
        "Define test closure and deliverables.",
        ["Final activities and documents after testing completion",
         "Includes finalising test reports and archiving assets",
         "Lessons learned documented",
         "Resources released formally",
         "Outstanding issues and recommendations recorded"],
        "Test closure includes finalising test reports, archiving test cases and results, analysing lessons learned, releasing resources, and documenting outstanding issues and recommendations. Key deliverables include test reports and metrics."
    ),
    (
        "What is the defect life cycle?",
        ["Progression of a defect from creation to closure",
         "States: New → Assigned → Open → Fixed → Verified → Closed",
         "Defects may be Reopened if fix is insufficient",
         "Tracked in defect management tools",
         "Provides accountability and audit trail"],
        "The defect life cycle includes states: New → Assigned → Open → Fixed → Test/Verified → Closed/Reopened, tracking defects from detection through resolution and verification."
    ),
    (
        "How do you estimate testing effort and resources?",
        ["Use historical data and expert judgment",
         "Apply three-point estimation for uncertainty",
         "Break down work using WBS",
         "Consider complexity, scope, and risk",
         "Review and adjust estimates iteratively"],
        "Test Managers estimate using techniques like expert judgment, analogous estimation, three-point estimation, work breakdown structure (WBS), considering complexity, scope, risk, and team capacity."
    ),
    (
        "Explain the concept of test data management.",
        ["Creating and maintaining realistic test data",
         "Ensures data is representative of production",
         "Complies with privacy regulations (GDPR, etc.)",
         "Maintains data integrity and consistency",
         "Includes data masking for sensitive information"],
        "Test data management involves identifying data requirements, creating test datasets that are realistic, secure, and representative, ensuring compliance with privacy regulations, and maintaining data integrity."
    ),
    (
        "What is a test automation framework?",
        ["Structured environment for automated test execution",
         "Includes tools, libraries, and guidelines",
         "Types: record/playback, keyword-driven, data-driven",
         "Improves test efficiency and reusability",
         "Requires investment in setup and maintenance"],
        "A test automation framework is a structured set of tools, libraries, and guidelines that provides a foundation for creating and executing automated tests efficiently, including record/playback, keyword-driven, or data-driven frameworks."
    ),
    (
        "Define test execution and monitoring.",
        ["Running test cases as planned and tracking progress",
         "Logging results and documenting defects",
         "Comparing actual vs planned progress",
         "Taking corrective actions when needed",
         "Providing regular status updates"],
        "Test execution involves running test cases as planned, logging results, documenting defects, and tracking metrics. Monitoring includes comparing actual progress against the test plan and taking corrective actions if needed."
    ),
    (
        "How do you handle testing in Agile environments?",
        ["Test continuously within each sprint",
         "Collaborate closely with developers",
         "Adapt plans based on sprint feedback",
         "Maintain minimal but sufficient documentation",
         "Focus on iterative delivery and quality"],
        "In Agile, Test Managers work in sprints, collaborate closely with developers, use continuous testing, maintain minimal documentation, adapt plans based on feedback, and focus on iterative delivery and quality."
    ),
    (
        "What is the difference between test strategy and test plan?",
        ["Strategy: long-term, organisation-wide approach",
         "Plan: project-specific, detailed execution document",
         "Strategy defines philosophy and methodology",
         "Plan defines scope, schedule, and resources",
         "Strategy is input to the test plan"],
        "Test strategy outlines the overall testing philosophy, methodology, and approach for the organisation, while a test plan is project-specific, detailing scope, resources, schedule, and approach for that project."
    ),
    (
        "Explain shift-left testing approach.",
        ["Moving testing earlier in the SDLC",
         "Includes requirements reviews and TDD",
         "Catches defects earlier at lower cost",
         "Promotes developer-tester collaboration",
         "Reduces rework and late-stage defects"],
        "Shift-left testing emphasises testing earlier in the development lifecycle, including requirements reviews, test-driven development (TDD), unit testing, and static analysis to catch defects early and reduce costs."
    ),
    (
        "What is a quality gate in testing?",
        ["Exit criteria checkpoint before proceeding to next phase",
         "Requires specific criteria to be met",
         "Prevents poor quality from advancing",
         "Examples: minimum coverage, zero critical defects",
         "Formally approved by stakeholders"],
        "A quality gate is a checkpoint where specific criteria must be met before proceeding to the next phase, such as achieving minimum test coverage, resolving critical defects, or meeting performance benchmarks."
    ),
    (
        "How do you handle schedule pressure and aggressive timelines?",
        ["Prioritise testing based on risk",
         "Use automation to increase efficiency",
         "Communicate constraints clearly to stakeholders",
         "Focus on high-risk, high-impact areas",
         "Escalate unrealistic expectations early"],
        "When facing schedule pressure, prioritise testing based on risk, use test automation for efficiency, communicate constraints clearly to stakeholders, focus on high-risk areas, and escalate unrealistic expectations early."
    ),
    (
        "Define Six Sigma principles in testing context.",
        ["Data-driven process improvement for quality",
         "Uses DMAIC: Define, Measure, Analyse, Improve, Control",
         "Targets reducing defects and variability",
         "Requires quantitative measurement",
         "Drives continuous improvement culture"],
        "Six Sigma in testing focuses on reducing defects and variability through data-driven analysis (DMAIC), process optimisation, and continuous improvement initiatives."
    ),
    (
        "What is the role of Test Manager in DevOps?",
        ["Ensure continuous quality in rapid deployment cycles",
         "Facilitate CI/CD pipeline integration",
         "Promote and govern test automation",
         "Maintain quality gates in pipelines",
         "Provide rapid feedback to development teams"],
        "In DevOps, Test Managers facilitate continuous integration/continuous deployment (CI/CD), promote test automation, maintain quality gates, provide rapid feedback, and ensure quality is maintained despite accelerated release cycles."
    ),
    (
        "Explain test case design techniques and their application.",
        ["BVA tests boundary input values",
         "EP groups inputs into equivalent classes",
         "State transition tests state-based behaviour",
         "Decision table covers complex logic combinations",
         "Cause-effect graphing maps inputs to outputs"],
        "Test case design techniques include boundary value analysis, equivalence partitioning, state transition testing, cause-effect graphing, and decision table testing—each targeting specific aspects of functionality systematically."
    ),
    (
        "How do you measure testing effectiveness and ROI?",
        ["Track defect detection rate and escape rate",
         "Calculate cost per defect found",
         "Measure schedule and effort variance",
         "ROI = (Benefits - Testing Cost) / Testing Cost",
         "Use trend analysis across projects"],
        "Testing effectiveness is measured through metrics: defect detection rate, defect escape rate, cost per defect, schedule variance, and ROI calculated as (Benefits - Testing Cost) / Testing Cost."
    ),
    (
        "What are current trends in software testing and quality assurance?",
        ["AI-powered test automation",
         "Machine learning for defect prediction",
         "Shift-left and continuous testing",
         "DevOps and CI/CD integration",
         "Quality as a business driver, not just a gate"],
        "Current trends include AI-powered test automation, machine learning for defect prediction, shift-left testing, DevOps integration, continuous testing, RPA in testing, and focus on quality as a business driver."
    ),
    (
        "What is SAP Test Manager (SAP TM)?",
        ["SAP module for test management in SAP implementations",
         "Provides centralised test management capabilities",
         "Integrates with SAP Solution Manager",
         "Supports automated and manual test execution",
         "Used for SAP upgrades, patches, and new implementations"],
        "SAP Test Manager is a comprehensive testing solution provided by SAP for managing and executing tests in SAP system implementations, upgrades, and patches. It integrates with other SAP tools and provides centralised test management capabilities."
    ),
    (
        "Explain the main features of SAP Test Manager.",
        ["Automated test execution and scheduling",
         "Test case and defect management",
         "Test scope configuration and coverage tracking",
         "Risk assessment and prioritisation",
         "Comprehensive reporting and dashboards"],
        "SAP TM features include automated test execution, test case management, defect management, test scope configuration, test data management, risk assessment, integration with SAP systems, and comprehensive reporting."
    ),
    (
        "What is the purpose of test scope in SAP Test Manager?",
        ["Defines what functionality needs to be tested",
         "Prevents testing gaps",
         "Organises test coverage by module or process",
         "Aligned with business blueprint",
         "Updated when scope changes occur"],
        "Test scope in SAP TM allows you to define and organise what functionality and areas of the SAP system need to be tested, helping to ensure comprehensive coverage and preventing testing gaps."
    ),
    (
        "Define test plan in SAP Test Manager.",
        ["Hierarchical structure organising test activities",
         "Contains test cycles, test sets, and test cases",
         "Provides structured approach to execution",
         "Tracks results at each level",
         "Created per project phase or release"],
        "In SAP TM, a test plan is a hierarchical structure that organises test cycles, test sets, test cases, and test steps, providing a structured approach to test execution and result tracking."
    ),
    (
        "What is a test cycle in SAP Test Manager?",
        ["Time-bound testing phase with specific objectives",
         "Can represent unit, system, or UAT phases",
         "Has defined scope, schedule, and deliverables",
         "Multiple cycles can run in parallel",
         "Progress tracked via cycle dashboards"],
        "A test cycle in SAP TM represents a specific phase of testing with defined objectives, scope, schedule, and deliverables. Multiple test cycles can be planned for different phases like unit testing, system testing, UAT, etc."
    ),
    (
        "Explain test sets in SAP Test Manager.",
        ["Logical grouping of related test cases",
         "Organised by functionality, module, or process",
         "Simplifies test assignment and execution",
         "Can be reused across test cycles",
         "Tracks execution status at group level"],
        "A test set in SAP TM is a collection of related test cases grouped together for execution, allowing testers to organise tests by functionality, module, or process flow for better management and execution."
    ),
    (
        "What is the role of a defect in SAP Test Manager?",
        ["Records and tracks identified issues",
         "Includes severity, priority, and steps to reproduce",
         "Tracks full defect lifecycle",
         "Linked to test cases and test cycles",
         "Enables defect trend analysis"],
        "A defect in SAP TM is a record of an issue found during testing, including details like severity, priority, steps to reproduce, expected vs actual results, and assignment. It tracks the defect lifecycle from creation to closure."
    ),
    (
        "Define SAP Solution Manager (SAP SolMan).",
        ["Centralised management tool for SAP systems",
         "Provides application lifecycle management",
         "Includes system monitoring and incident management",
         "Integrates with SAP Test Manager",
         "Foundation for SAP quality management"],
        "SAP Solution Manager is an SAP product that provides comprehensive management capabilities for SAP solutions, including application lifecycle management, system monitoring, incident management, and integration with SAP Test Manager."
    ),
    (
        "How does SAP Test Manager integrate with SAP Solution Manager?",
        ["Provides centralised project and test management",
         "Seamless defect tracking across both tools",
         "Unified test execution monitoring",
         "Shared reporting and dashboards",
         "Enables traceability from requirement to test result"],
        "SAP TM integrates with SAP SolMan to provide centralised project management, defect tracking, test execution monitoring, and reporting, allowing seamless communication between testing and project management teams."
    ),
    (
        "What is the ASAP methodology in SAP context?",
        ["Accelerated SAP implementation framework",
         "Five phases: Project Preparation, Blueprint, Realization, Final Prep, Go-Live",
         "Testing embedded in each phase",
         "SAP TM supports test activities throughout",
         "Proven methodology for SAP implementations"],
        "ASAP (Accelerated SAP) is SAP's proven implementation methodology comprising five phases: Project Preparation, Business Blueprint, Realization, Final Preparation, and Go-Live. Testing is embedded in each phase with SAP TM supporting test activities."
    ),
    (
        "Explain quality records in SAP Test Manager.",
        ["Documents proving test execution and results",
         "Includes test logs, screenshots, and defect reports",
         "Serve as audit trail and compliance evidence",
         "Required for formal sign-off",
         "Stored in SAP TM for traceability"],
        "Quality records in SAP TM are documents that provide evidence of test execution, including test logs, screenshots, defect reports, and sign-offs, serving as audit trail and quality documentation."
    ),
    (
        "What is a template in SAP Test Manager?",
        ["Reusable structure for test cases and configurations",
         "Ensures consistency across test cases",
         "Reduces effort for similar processes",
         "Can be shared across projects",
         "Updated centrally when processes change"],
        "Templates in SAP TM allow you to create reusable test case structures, defect templates, and configuration templates, improving consistency, reducing effort, and speeding up test creation for similar processes."
    ),
    (
        "Define test data configuration in SAP Test Manager.",
        ["Setting up test data for test execution",
         "Includes master data and transactional data",
         "Simulates real business scenarios",
         "Organised per test cycle or test set",
         "Critical for accurate test results"],
        "Test data configuration in SAP TM involves identifying, preparing, and organising data needed for testing, including creation of test datasets, master data, and transactional data that simulates real business scenarios."
    ),
    (
        "What is the purpose of the roadmap in SAP Test Manager?",
        ["Strategic planning for testing activities",
         "High-level view of testing phases and milestones",
         "Helps stakeholders understand testing approach",
         "Aligned with project timeline",
         "Updated as project evolves"],
        "The roadmap in SAP TM provides a high-level view of testing strategy, phases, milestones, and deliverables, helping stakeholders understand the overall testing approach and timeline for a project or implementation."
    ),
    (
        "Explain function module testing in SAP context.",
        ["Testing individual SAP function modules",
         "Verifies reusable ABAP code units work correctly",
         "Performed before integration testing",
         "Uses test driver programs or SE37",
         "Foundation for higher-level testing"],
        "Function module testing in SAP involves testing individual function modules (reusable code units in SAP ABAP) to verify they work correctly before they're integrated into larger processes and applications."
    ),
    (
        "What is the SAP Testing Pyramid?",
        ["Hierarchical approach to test coverage in SAP",
         "More unit tests at the base",
         "Fewer integration tests in the middle",
         "Fewer end-to-end tests at the top",
         "Optimises efficiency and maintainability"],
        "The SAP Testing Pyramid recommends more unit tests at the base, fewer integration tests in the middle, and fewer end-to-end tests at the top, optimising test efficiency, speed, and maintainability."
    ),
    (
        "Define RFC (Remote Function Call) testing in SAP.",
        ["Testing external system integration via RFC",
         "Verifies data transfer between SAP and external systems",
         "Ensures correct process flow across system boundaries",
         "Includes synchronous and asynchronous RFCs",
         "Critical for integrated landscapes"],
        "RFC testing in SAP involves testing remote function calls that enable communication between SAP and external systems or between SAP modules, ensuring data transfer and process flow work correctly."
    ),
    (
        "What is Business Process testing in SAP?",
        ["Testing complete business processes end-to-end",
         "Validates all integrated steps work together",
         "Examples: procure-to-pay, order-to-cash",
         "Involves multiple SAP modules",
         "Closest to real business scenarios"],
        "Business Process testing in SAP TM involves testing complete business processes from start to finish (e.g., procure-to-pay, order-to-cash), ensuring all integrated steps work together seamlessly."
    ),
    (
        "Explain SAP process variants and testing implications.",
        ["Different configurations of the same SAP process",
         "Based on company-specific customisation",
         "Each variant must be tested independently",
         "Increases scope of regression testing",
         "Documented in business blueprint"],
        "SAP process variants represent different ways business processes can be configured in SAP based on company needs. Testing must cover all configured variants to ensure each functions correctly in the implemented environment."
    ),
    (
        "What is impact analysis in SAP testing?",
        ["Assessing effects of changes on other processes",
         "Identifies additional areas requiring regression testing",
         "Reduces risk of unintended consequences",
         "Performed before each change or patch",
         "Supported by SAP TM scope management"],
        "Impact analysis in SAP TM helps identify which business processes and functionalities are affected by changes or patches, determining what additional testing is needed beyond the direct change area."
    ),
    (
        "Define SAP Maintenance Update (SAP Notes) testing.",
        ["Testing SAP system patches and corrections",
         "Ensures fixes work correctly post-application",
         "Validates no new issues introduced",
         "Includes regression testing of affected areas",
         "Documented and tracked in SAP TM"],
        "When SAP releases maintenance updates or notes (corrections), testing ensures the fixes work correctly, don't introduce new issues, and don't negatively impact existing functionality through regression testing."
    ),
    (
        "What is the role of Risk Assessment in SAP testing?",
        ["Identifying and prioritising testing areas based on risk",
         "Considers complexity, criticality, and change history",
         "Guides test effort allocation",
         "Mitigates business risk from system changes",
         "Reviewed and updated throughout the project"],
        "Risk assessment in SAP TM involves identifying high-risk areas (complex processes, critical functionalities, areas with history of issues) and prioritising testing efforts to mitigate business risks."
    ),
    (
        "Explain golden record testing in SAP context.",
        ["Testing with representative real business data",
         "Provides accurate, real-world test results",
         "Uses actual production data scenarios",
         "Critical for go-live readiness validation",
         "Performed during Final Preparation phase"],
        "Golden record testing uses representative, realistic business data and scenarios from actual operations, providing more accurate testing results that reflect real-world conditions in the SAP system."
    ),
    (
        "What is Cut-Over testing in SAP implementations?",
        ["Final validation before going live",
         "Uses production-like data and scenarios",
         "Validates configuration, migration, and interfaces",
         "Performed in Final Preparation phase",
         "Represents the last quality gate before go-live"],
        "Cut-over testing (or golden record testing) is performed immediately before go-live to validate all configurations, data migration, and interfaces work correctly using production-like scenarios and actual data."
    ),
    (
        "Define SAP Security testing.",
        ["Testing user authorisations and data security",
         "Verifies role-based access controls (RBAC)",
         "Ensures sensitive data is protected",
         "Tests segregation of duties (SoD)",
         "Compliance requirement for most organisations"],
        "SAP Security testing verifies that user roles and authorisations are configured correctly, data access is properly restricted, and sensitive information is protected through role-based access controls (RBAC)."
    ),
    (
        "What is SLA (Service Level Agreement) testing in SAP?",
        ["Testing system performance against agreed standards",
         "Validates response time, availability, and throughput",
         "Includes batch processing time validation",
         "Benchmarked against contractual SLAs",
         "Results reported to management for sign-off"],
        "SLA testing in SAP ensures the system meets agreed service levels like response time, availability, throughput, and batch processing times, validating that system performance meets business requirements."
    ),
    (
        "Explain batch job testing in SAP.",
        ["Testing scheduled background processes",
         "Verifies jobs execute correctly on schedule",
         "Checks output accuracy and completeness",
         "Validates completion within acceptable timeframes",
         "Critical for payroll, reporting, and data archival"],
        "Batch job testing involves testing SAP background jobs that run on schedule (like payroll, reporting, data archival) to ensure they execute correctly, produce accurate results, and complete within acceptable timeframes."
    ),
    (
        "What is interface testing in SAP?",
        ["Testing data flow between SAP and external systems",
         "Verifies data accuracy and completeness",
         "Tests via RFC, SOAP, REST protocols",
         "Ensures no data loss during transmission",
         "Covers both inbound and outbound interfaces"],
        "Interface testing in SAP TM verifies that data transfers correctly between SAP and integrated systems (e.g., via RFC, SOAP, REST), ensuring data accuracy, completeness, and no loss during transmission."
    ),
    (
        "Define master data testing in SAP.",
        ["Testing creation and management of reference data",
         "Covers customers, materials, vendors, GL accounts",
         "Validates creation rules and referential integrity",
         "Checks data replication across systems",
         "Foundation for transactional testing"],
        "Master data testing verifies that master data (customers, materials, vendors, GL accounts) is created, updated, and maintained correctly in SAP, including validation rules and referential integrity."
    ),
    (
        "What is transactional testing in SAP?",
        ["Testing business transactions from creation to completion",
         "Validates all processing steps and system updates",
         "Examples: sales orders, purchase orders, invoices",
         "Checks posting, accounting entries, and workflow",
         "Core of SAP functional testing"],
        "Transactional testing in SAP involves testing complete business transactions (sales orders, purchase orders, invoices) from creation through processing to closure, validating all steps and system updates."
    ),
    (
        "Explain configuration testing in SAP.",
        ["Testing system customisation and IMG settings",
         "Verifies configuration meets business requirements",
         "Includes Customising and enterprise structure settings",
         "Performed after each configuration change",
         "Ensures best practice alignment"],
        "Configuration testing verifies that SAP configuration parameters, Customising IMG settings, and system customisation have been set up correctly according to business requirements and best practices."
    ),
    (
        "What is the purpose of test reporting in SAP Test Manager?",
        ["Communicating test progress and quality metrics",
         "Provides visibility to stakeholders",
         "Tracks defect status and test coverage",
         "Supports go/no-go decision making",
         "Includes dashboards and detailed reports"],
        "Test reporting in SAP TM provides stakeholders with visibility into test execution progress, defect status, quality metrics, test coverage, and recommendations through various reports and dashboards."
    ),
    (
        "Define baseline testing in SAP.",
        ["Establishing a reference point for future comparisons",
         "Measures current system functionality and performance",
         "Enables regression detection over time",
         "Supports quality trending across releases",
         "Performed at project start and after major changes"],
        "Baseline testing in SAP TM establishes a reference point for system functionality, performance, and defect levels, allowing future test results to be compared for regression detection and quality trending."
    ),
    (
        "What is Compatibility testing in SAP context?",
        ["Testing across SAP releases and database platforms",
         "Covers Oracle, SQL Server, and DB2 databases",
         "Validates across operating systems and browsers",
         "Ensures upgrade compatibility",
         "Critical for landscape standardisation"],
        "Compatibility testing in SAP ensures applications work correctly across different SAP versions, database platforms (Oracle, SQL Server, DB2), operating systems, and browser combinations."
    ),
    (
        "Explain the role of test leads in SAP testing.",
        ["Supervise test team and manage daily execution",
         "Assign work to testers and track progress",
         "Resolve execution blockers and escalate issues",
         "Maintain quality standards within the team",
         "Provide status updates to the Test Manager"],
        "Test leads in SAP projects manage daily test execution, assign work to testers, track progress, resolve issues, escalate blockers, maintain quality standards, and provide status updates to management."
    ),
    (
        "What is knowledge transfer in SAP testing context?",
        ["Training support team on new processes",
         "Documenting test results and lessons learned",
         "Creating support runbooks for post-go-live",
         "Ensures continuity after project team leaves",
         "Critical for long-term system stability"],
        "Knowledge transfer in SAP testing includes training the support team on new processes, documentation of test results, lessons learned, and creation of support runbooks for post-go-live operations."
    ),
    (
        "Define hyper-care period in SAP implementations.",
        ["Intensive support phase immediately after go-live",
         "Extended support teams on-site or on-call",
         "Monitors system performance and stability",
         "Addresses issues immediately",
         "Ensures business continuity during transition"],
        "The hyper-care period is the intensive support phase following SAP go-live where extended support teams are on-site to monitor system performance, address issues immediately, and ensure business continuity."
    ),
    (
        "What is the role of test metrics in SAP projects?",
        ["Monitoring and controlling test quality and progress",
         "Tracks defect density and detection rate",
         "Monitors test execution rate and coverage",
         "Supports data-driven decision-making",
         "Reported to project steering committee"],
        "Test metrics in SAP projects track indicators like defect density, test execution rate, defect detection rate, and test coverage, enabling data-driven decision-making and quality assessment."
    ),
    (
        "Explain value stream mapping in SAP context.",
        ["Visualising business processes to identify inefficiencies",
         "Maps current vs future state processes",
         "Identifies redundancies and improvement opportunities",
         "Ensures SAP configuration supports optimised processes",
         "Input to test scope and business process design"],
        "Value stream mapping in SAP testing involves mapping current business processes to identify inefficiencies, redundancies, and improvement opportunities, ensuring SAP configuration supports optimised processes."
    ),
    (
        "What is the importance of sign-off in SAP testing?",
        ["Formal approval and acceptance of test results",
         "Authorises progression to next phase or go-live",
         "Documents stakeholder agreement on quality",
         "Provides legal and audit protection",
         "Required from business, IT, and management"],
        "Sign-off in SAP testing represents formal approval from stakeholders, business users, and project management that testing is complete, quality is acceptable, and the system is ready for the next phase or go-live."
    ),
    (
        "Define regression pack in SAP Test Manager.",
        ["Collection of critical test cases for regression testing",
         "Covers core business functionality",
         "Executed after each change or patch",
         "Maintained and updated over time",
         "Reduces effort for recurring regression cycles"],
        "A regression pack in SAP TM is a curated set of critical test cases covering core business functionality that are executed regularly to ensure system changes haven't broken existing functionality."
    ),
    (
        "What is the role of Test Manager in SAP Go-Live?",
        ["Ensure all testing is complete and sign-offs obtained",
         "Manage final cut-over testing",
         "Confirm test infrastructure is ready",
         "Coordinate with support and hypercare teams",
         "Manage production issues during go-live window"],
        "The Test Manager during SAP go-live ensures all testing is complete, critical issues are resolved, test infrastructure is ready, supports cut-over testing, manages any production issues, and coordinates with support teams."
    ),
    (
        "Explain P-Fit (Process Fit) in SAP implementations.",
        ["Evaluating how well customer processes fit SAP standard",
         "Identifies gaps requiring customisation",
         "Performed during Business Blueprint phase",
         "Impacts testing scope and complexity",
         "Higher fit = less customisation = simpler testing"],
        "P-Fit is an SAP evaluation of how well customer processes fit with SAP standard functionality, used during design to identify gaps requiring customisation and impacts on testing scope."
    ),
    (
        "What is the importance of Prerequisites in test cases?",
        ["Establishing conditions before test execution",
         "Ensures tests start from a known, valid state",
         "Includes data, system settings, and user access",
         "Prevents false failures due to missing setup",
         "Documented clearly for all testers"],
        "Prerequisites in SAP test cases define the required state of data, system settings, and user access needed before a test can be executed, ensuring tests start from a known, valid state."
    ),
    (
        "Define end-user testing (EUT) in SAP context.",
        ["Testing by actual business users",
         "Validates system meets real user needs",
         "Assesses usability and intuitiveness",
         "Confirms alignment with business processes",
         "Final validation before go-live decision"],
        "End-user testing in SAP involves actual business users executing test cases to validate that the system meets their needs, works intuitively, and aligns with their business processes before go-live."
    ),
    (
        "What is the role of Test Manager in quality assurance?",
        ["Ensure testing meets quality standards and objectives",
         "Follows established procedures and methodologies",
         "Delivers comprehensive test coverage",
         "Identifies defects early to reduce cost",
         "Supports informed go/no-go decisions"],
        "The Test Manager ensures that testing activities follow established procedures, meet quality standards, deliver comprehensive coverage, identify defects early, and support informed go/no-go decisions."
    ),
    (
        "Explain defect triage in SAP testing.",
        ["Reviewing and prioritising defects for resolution",
         "Assesses severity and business impact",
         "Assigns defects to responsible teams",
         "Determines if workarounds are acceptable",
         "Daily or regular triage meetings during execution"],
        "Defect triage in SAP testing involves reviewing identified defects, assessing severity and business impact, prioritising fixes, assigning responsibility, and determining if workarounds are acceptable."
    ),
    (
        "What is the importance of test documentation in SAP?",
        ["Creates audit trail and knowledge base",
         "Records what was tested and results obtained",
         "Documents decisions and sign-offs",
         "Serves as compliance and regulatory evidence",
         "Knowledge resource for future projects"],
        "Test documentation in SAP projects creates a complete record of what was tested, how it was tested, results obtained, and decisions made, serving as audit evidence and knowledge resource for future projects."
    ),
    (
        "Define continuous improvement in SAP testing.",
        ["Ongoing process optimisation based on lessons learned",
         "Captures learnings from each testing cycle",
         "Analyses defect trends for root causes",
         "Implements enhancements for future projects",
         "Builds institutional testing knowledge"],
        "Continuous improvement in SAP testing involves capturing lessons learned from each testing cycle, analysing defect trends, identifying process improvements, and implementing enhancements for future projects."
    ),
    (
        "What is the strategic importance of Test Manager in SAP implementations?",
        ["Driving quality and managing risk across the programme",
         "Ensuring comprehensive testing and business coverage",
         "Maintaining timeline adherence under pressure",
         "Reducing post-go-live issues and business disruption",
         "Supporting business value realisation from SAP investment"],
        "The Test Manager in SAP implementations is critical to project success by ensuring comprehensive testing, managing quality risks, maintaining timeline adherence, reducing post-go-live issues, and supporting business value realisation."
    ),
]

# ── 100 Scrum Master Q&A ──────────────────────────────────────────────────
SCRUM_MASTER_QA = [
    (
        "What is Scrum?",
        ["Framework for managing product development iteratively",
         "Uses fixed-length cycles called sprints",
         "Defines roles: Product Owner, Scrum Master, Development Team",
         "Artifacts: Product Backlog, Sprint Backlog, Increment",
         "Ceremonies: Sprint Planning, Daily Standup, Review, Retrospective"],
        "Scrum is an agile framework for managing product development through iterative cycles called sprints. It emphasizes teamwork, accountability, and continuous improvement. Scrum uses defined roles, artifacts, and ceremonies to deliver working software incrementally."
    ),
    (
        "What are the three pillars of Scrum?",
        ["Transparency — all aspects visible to stakeholders",
         "Inspection — regular examination of artifacts and progress",
         "Adaptation — adjusting process based on inspection findings",
         "Together they form the empirical process control foundation",
         "All ceremonies embody these three pillars"],
        "The three pillars of Scrum form the foundation: Transparency ensures all aspects of the process are visible; Inspection involves regular examination of artifacts and progress; Adaptation means adjusting based on inspection findings to minimise problems."
    ),
    (
        "Define the role of Scrum Master.",
        ["Servant-leader ensuring team follows Scrum practices",
         "Facilitates all Scrum ceremonies",
         "Removes impediments blocking the team",
         "Coaches team on self-organisation",
         "Protects team from external disruptions"],
        "A Scrum Master is a servant-leader who ensures the team understands and follows Scrum principles. They facilitate ceremonies, remove impediments, coach the team on self-organisation, protect the team from disruptions, and foster continuous improvement."
    ),
    (
        "What is the difference between Scrum Master and Project Manager?",
        ["Scrum Master: servant-leader enabling the team",
         "Project Manager: command-and-control directing work",
         "Scrum Master focuses on process and team enablement",
         "Project Manager focuses on scope, schedule, and resources",
         "Scrum Master has no authority to assign tasks"],
        "A Scrum Master is a servant-leader who enables the team and removes obstacles, while a Project Manager traditionally controls, directs, and commands. Scrum Masters focus on process and enablement; Project Managers focus on scope, schedule, and resources."
    ),
    (
        "Explain the Product Owner role in Scrum.",
        ["Manages and prioritises the Product Backlog",
         "Defines acceptance criteria for user stories",
         "Represents customer and business interests",
         "Accepts or rejects completed work",
         "Maximises return on investment"],
        "The Product Owner manages and prioritises the Product Backlog, defines acceptance criteria, communicates vision to the team, accepts completed work, and ensures the team builds the right product. They represent customer interests and maximise ROI."
    ),
    (
        "What is a Development Team in Scrum?",
        ["Self-organising group creating product increments",
         "Cross-functional — all skills needed to deliver value",
         "Typically 5–9 members",
         "Commits to sprint goals collectively",
         "No sub-teams or hierarchy within the team"],
        "A Development Team in Scrum is cross-functional, self-organising, and responsible for creating working product increments. Typically 5–9 members with all skills needed to deliver value, they commit to sprint goals and self-manage their work."
    ),
    (
        "Define sprint in Scrum.",
        ["Fixed, repeating time-box for completing work",
         "Typically 1–4 weeks in duration",
         "Results in a potentially shippable increment",
         "Sprint goal set during Sprint Planning",
         "Cannot be shortened or lengthened mid-sprint"],
        "A sprint is a fixed, repeating time-box (typically 1–4 weeks) during which the team works to complete a set of user stories from the Product Backlog. Each sprint results in a potentially shippable product increment."
    ),
    (
        "What is the Product Backlog?",
        ["Ordered list of all features and requirements",
         "Maintained and prioritised by the Product Owner",
         "Dynamic — continuously refined and updated",
         "Items prioritised by value, dependencies, and risk",
         "Single source of truth for product work"],
        "The Product Backlog is an ordered list of all features, enhancements, fixes, and requirements for the product, maintained by the Product Owner. It's dynamic and continuously refined based on feedback and changing business needs."
    ),
    (
        "Define Sprint Backlog.",
        ["Committed work items selected for the current sprint",
         "Includes plan for delivering each item",
         "Contains tasks, technical work, and acceptance criteria",
         "Only the Development Team can modify it during sprint",
         "Visible on the Scrum board"],
        "The Sprint Backlog is the list of Product Backlog items selected for the current sprint, along with the plan for delivering them. Only the Development Team can modify the Sprint Backlog during the sprint."
    ),
    (
        "What is a User Story?",
        ["Short description of a feature from user perspective",
         "Format: As a [user], I want [feature], so that [benefit]",
         "Captures customer value",
         "Must be testable and fit within a sprint",
         "Includes acceptance criteria"],
        "A User Story is a short, simple description of a feature written from a user's perspective. It captures customer value, is testable, and fits within a sprint. Acceptance criteria define when the story is complete."
    ),
    (
        "Explain story points.",
        ["Relative estimation unit for effort and complexity",
         "Common scales: Fibonacci (1, 2, 3, 5, 8, 13…)",
         "Enables velocity tracking",
         "Relative — compared to other stories, not time",
         "Whole team participates in estimation (Planning Poker)"],
        "Story points are units used to estimate the relative size, effort, and complexity of User Stories. Common scales are Fibonacci or T-shirt sizes. They enable velocity tracking and sprint planning."
    ),
    (
        "What is velocity in Scrum?",
        ["Total story points completed in a sprint",
         "Helps team understand capacity",
         "Used to predict future sprint commitments",
         "Typically stabilises after 3–4 sprints",
         "Input to release planning"],
        "Velocity is the total number of story points completed in a sprint. It helps the team understand their capacity, predict future commitments, identify trends, and improve planning accuracy."
    ),
    (
        "Define Sprint Planning ceremony.",
        ["Team commits to sprint work and creates Sprint Backlog",
         "Product Owner presents prioritised items",
         "Team discusses, estimates, and commits",
         "Time-boxed: 4 hours for 2-week sprint",
         "Produces sprint goal and Sprint Backlog"],
        "Sprint Planning is a time-boxed ceremony where the Product Owner presents prioritised items, the team discusses and estimates them, and commits to work they believe they can complete. The Sprint Backlog is created during this ceremony."
    ),
    (
        "What is the Daily Standup ceremony?",
        ["15-minute daily synchronisation meeting",
         "Each member shares: done, doing, blockers",
         "Peer-to-peer — not a status report to Scrum Master",
         "Held at same time and place each day",
         "Identifies impediments early"],
        "The Daily Standup is a 15-minute time-boxed ceremony where each team member shares what they completed yesterday, what they'll do today, and any impediments. It provides visibility, improves coordination, and identifies blockers early."
    ),
    (
        "Explain Sprint Review ceremony.",
        ["Demonstrates completed work to stakeholders",
         "Gathers feedback on the increment",
         "Product Owner discusses backlog priorities",
         "Time-boxed: 1–2 hours for 2-week sprint",
         "Drives next sprint direction"],
        "The Sprint Review is held at the end of the sprint where the team demonstrates completed work to stakeholders, gathers feedback, and discusses Product Backlog priorities with the Product Owner."
    ),
    (
        "What is Sprint Retrospective ceremony?",
        ["Team reflects on process and commits to improvements",
         "Covers what went well, what didn't, what to improve",
         "Time-boxed: 1–1.5 hours for 2-week sprint",
         "Held after Sprint Review",
         "Action items carried into next sprint"],
        "The Sprint Retrospective is held after the Sprint Review where the team reflects on what went well, what didn't, and what they can improve. They commit to specific process improvements for the next sprint."
    ),
    (
        "Define definition of done (DoD).",
        ["Shared understanding of what completed means",
         "Includes: code reviewed, tests passed, documentation updated",
         "Prevents incomplete work from accumulating",
         "Applied consistently to every story",
         "Collaboratively created by the whole team"],
        "Definition of Done is a shared understanding of what completed means. It includes criteria like code review done, tests written and passed, documentation updated, and meets acceptance criteria. DoD prevents incomplete work from accumulating."
    ),
    (
        "What is an impediment in Scrum?",
        ["Obstacle blocking team's progress",
         "Can be technical, organisational, or external",
         "Identified in Daily Standup",
         "Scrum Master owns removal",
         "Tracked until resolved"],
        "An impediment is anything that prevents the team from completing their work or achieving sprint goals, such as blocked dependencies, missing resources, or external blockers. Removing impediments is a core Scrum Master responsibility."
    ),
    (
        "Explain backlog refinement in Scrum.",
        ["Ongoing activity to prepare upcoming backlog items",
         "Product Owner and team collaborate",
         "Items reviewed, clarified, estimated, and prioritised",
         "Typically 10% of team's sprint capacity",
         "Ensures stories are sprint-ready"],
        "Backlog refinement is an ongoing activity where the Product Owner and team collaboratively review, clarify, estimate, and prioritise upcoming Product Backlog items to prepare them for future sprints."
    ),
    (
        "What is the purpose of the Definition of Ready (DoR)?",
        ["Criteria items must meet before entering a sprint",
         "Requirements must be clear and understood",
         "Acceptance criteria must be defined",
         "Story must be estimated",
         "No external blocking dependencies"],
        "Definition of Ready establishes criteria that Product Backlog items must meet before being pulled into a sprint, such as clear requirements, acceptance criteria defined, estimated, and no blocking dependencies."
    ),
    (
        "Define burndown chart.",
        ["Visual representation of remaining work over time",
         "Vertical axis: remaining story points or hours",
         "Horizontal axis: sprint days",
         "Ideal line shows expected linear progress",
         "Reveals capacity and scope change issues"],
        "A burndown chart displays the amount of work remaining against time. An ideal burndown shows remaining work decreasing consistently. Actual burndown may vary, revealing capacity, scope changes, and process issues."
    ),
    (
        "What is a burnup chart?",
        ["Visual representation of completed work over time",
         "Shows progress toward the release goal",
         "Clearly shows scope increases",
         "Less demoralising than burndown",
         "Useful for stakeholder communication"],
        "A burnup chart displays completed work against time. It shows progress toward the goal, clearly shows scope increases, and is useful for stakeholder communication."
    ),
    (
        "Explain velocity chart.",
        ["Tracks story points completed per sprint over time",
         "Reveals team capacity trends",
         "Helps with future sprint planning",
         "Identifies capacity changes or process issues",
         "Stable velocity indicates predictable performance"],
        "A velocity chart shows story points completed in each sprint over time. It reveals team capacity trends, helps with future sprint planning, and identifies process improvements or capacity changes."
    ),
    (
        "What is a task in Scrum?",
        ["Unit of work contributing to completing a User Story",
         "Created during Sprint Planning",
         "Estimated in hours (typically 4–16 hours)",
         "Tracked on the Sprint Backlog",
         "More concrete and smaller than User Stories"],
        "A task is a unit of work that contributes to completing a User Story. Tasks are created during Sprint Planning, estimated in hours, and tracked on the Sprint Backlog."
    ),
    (
        "Define capacity planning in Scrum.",
        ["Estimating available team hours for a sprint",
         "Accounts for vacations, meetings, and other commitments",
         "Ensures realistic sprint goals",
         "Calculated per team member per sprint",
         "Input to Sprint Planning"],
        "Capacity planning involves calculating available team hours for a sprint by accounting for planned absences, meetings, and other commitments. This helps the team commit to realistic sprint goals."
    ),
    (
        "What is a spike in Scrum?",
        ["Time-boxed investigation to reduce uncertainty",
         "Typically 1–2 days duration",
         "Produces knowledge, not product increment",
         "Helps with more accurate future estimation",
         "Used for complex or unfamiliar problems"],
        "A spike is a short time-boxed investigation to reduce technical uncertainty, understand complex problems, or explore solutions. Spikes produce knowledge, not product increment, and help with more accurate future estimation."
    ),
    (
        "Explain technical debt.",
        ["Cost of deferred quality improvements and shortcuts",
         "Accumulates when speed is prioritised over quality",
         "Reduces long-term team productivity",
         "Must be balanced with feature development",
         "Scrum Master ensures team dedicates time to paying it down"],
        "Technical debt refers to deferred quality improvements, shortcuts taken, or inadequate refactoring that creates future costs. Scrum Masters should ensure the team balances adding features with managing technical debt."
    ),
    (
        "What is a release in Scrum?",
        ["Collection of sprints delivering significant value",
         "Subset of features selected for production deployment",
         "Planned collaboratively with Product Owner",
         "May span multiple sprints",
         "Based on product roadmap priorities"],
        "A release is a subset of features selected to be deployed to production, typically comprising multiple sprints. Release planning involves deciding which Product Backlog items will be included to deliver business value."
    ),
    (
        "Define releasable increment.",
        ["Potentially shippable product piece from a sprint",
         "Meets the Definition of Done",
         "Ready for release to production if needed",
         "Inspectable by stakeholders",
         "Accumulates toward full release readiness"],
        "A releasable increment is a piece of the product that meets the Definition of Done and is potentially ready for release to production. Each sprint produces an increment that could be released."
    ),
    (
        "What is the role of Scrum Master in impediment removal?",
        ["Identifies impediments through standups and observation",
         "Resolves directly when within team's control",
         "Escalates to management when organisational",
         "Facilitates team problem-solving",
         "Tracks all impediments until resolved"],
        "The Scrum Master identifies impediments through daily standups and team interactions, resolves them by working with relevant parties, escalating to management when needed, or facilitating team problem-solving."
    ),
    (
        "Explain servant leadership in Scrum context.",
        ["Leading by enabling and supporting others",
         "Primary focus: serving the team's needs",
         "Facilitates rather than directs",
         "Removes barriers for team success",
         "Contrasts with command-and-control leadership"],
        "Servant leadership means the Scrum Master's primary focus is serving the team's needs, enabling their success, and supporting their growth. Rather than directing, they facilitate, coach, and remove barriers."
    ),
    (
        "What is a self-organising team in Scrum?",
        ["Team decides how to accomplish sprint goals",
         "No external direction on how to do the work",
         "Determines task allocation and work methods",
         "Requires psychological safety to function",
         "Scrum Master enables by removing obstacles"],
        "A self-organising team determines how they'll accomplish sprint goals without external direction. They decide task allocation, work methods, and problem-solving approaches. The Scrum Master enables self-organisation by removing obstacles."
    ),
    (
        "Define cross-functional team.",
        ["Team with all skills needed to deliver the product",
         "No dependency on external specialists for delivery",
         "Includes development, testing, design skills",
         "Faster and more autonomous",
         "Reduces handoffs and waiting time"],
        "A cross-functional team possesses all the skills needed to deliver a working product increment without depending on external specialists. This includes development, testing, documentation, and design skills."
    ),
    (
        "What is stakeholder engagement in Scrum?",
        ["Involving interested parties in product development",
         "Sprint Reviews are primary engagement mechanism",
         "Scrum Master facilitates engagement",
         "Manages stakeholder expectations",
         "Ensures stakeholders see value delivery"],
        "Stakeholder engagement in Scrum involves Sprint Reviews, feedback sessions, and regular communication about progress. The Scrum Master facilitates this engagement, manages expectations, and ensures stakeholders see value delivery."
    ),
    (
        "Explain the concept of done increments.",
        ["Tangible, working product pieces completed each sprint",
         "Meet the Definition of Done",
         "Inspectable and create feedback opportunities",
         "Enable visible progress measurement",
         "Accumulate toward release readiness"],
        "Done increments represent working, potentially releasable product pieces completed each sprint. They're inspectable, create feedback opportunities, and enable the team to see visible progress."
    ),
    (
        "What is a Scrum board?",
        ["Visual representation of sprint work status",
         "Columns: To Do, In Progress, In Review, Done",
         "Provides transparency on work status",
         "Identifies bottlenecks at a glance",
         "Tools: Jira, Azure DevOps, Trello, physical board"],
        "A Scrum board (physical or digital) displays sprint work items with status columns. It provides transparency on work status, identifies bottlenecks, and facilitates team coordination."
    ),
    (
        "Define continuous integration.",
        ["Regularly merging code into shared repository",
         "Triggers automated builds and tests",
         "Catches integration issues early",
         "Reduces bugs and rework",
         "Scrum Masters should promote CI practices"],
        "Continuous Integration is a practice where developers frequently merge code into a shared repository, triggering automated builds and tests. CI catches integration issues early, reduces bugs, and enables faster feedback."
    ),
    (
        "What is continuous deployment?",
        ["Automatically releasing tested code to production",
         "Code passing automated tests deployed automatically",
         "Enables rapid, frequent releases",
         "Reduces manual deployment effort",
         "Scrum Masters support CD in their teams"],
        "Continuous Deployment is the practice of automatically deploying code that passes automated tests to production. It enables rapid, frequent releases, faster feedback, and reduced manual effort."
    ),
    (
        "Explain pair programming.",
        ["Two developers writing code collaboratively",
         "Driver writes code; navigator reviews in real time",
         "Improves code quality and knowledge sharing",
         "Reduces bugs and defects",
         "Enhances team communication"],
        "Pair programming involves two developers working together at one computer—one writes code (driver) while the other reviews (navigator). Benefits include improved code quality, knowledge sharing, fewer bugs, and enhanced communication."
    ),
    (
        "What is test-driven development (TDD)?",
        ["Write automated tests before writing implementation code",
         "Red-Green-Refactor cycle",
         "Ensures code is testable and well-designed",
         "Produces fewer defects",
         "Scrum Masters should encourage TDD adoption"],
        "Test-Driven Development is a practice where developers write automated tests before implementing features. This ensures code is testable, well-designed, and produces fewer defects."
    ),
    (
        "Define code review.",
        ["Peer examination of code changes before merging",
         "Identifies issues and ensures standards compliance",
         "Shares knowledge across the team",
         "Improves overall code quality",
         "Part of the Definition of Done"],
        "Code review involves peers examining code changes before merging to identify issues, share knowledge, ensure standards compliance, and improve overall code quality. It's a critical quality practice Scrum Masters should ensure the team follows."
    ),
    (
        "What is retrospective action items?",
        ["Specific improvements committed at end of retrospective",
         "Must be SMART: Specific, Measurable, Achievable, Relevant, Time-bound",
         "Carried into next sprint as tasks",
         "Reviewed at start of next retrospective",
         "Scrum Master tracks completion"],
        "Retrospective action items are specific improvements the team commits to implementing in the next sprint. Scrum Masters ensure action items are SMART and tracked to completion."
    ),
    (
        "Explain the concept of inspect and adapt.",
        ["Core Scrum principle of continuous improvement",
         "Regular inspection of process, product, and progress",
         "Adaptation based on inspection findings",
         "Embodied in Sprint Reviews and Retrospectives",
         "Drives iterative improvement over time"],
        "Inspect and Adapt is a core Scrum principle where the team regularly inspects their process, product, and progress, then adapts based on findings. This happens in Sprint Reviews and Retrospectives, driving continuous improvement."
    ),
    (
        "What is a blocker in agile terms?",
        ["Obstacle preventing task completion",
         "Identified in Daily Standup",
         "Scrum Master owns removal",
         "Tracked until resolved",
         "Can be technical, organisational, or external"],
        "A blocker is anything preventing a team member from completing their work, such as missing information, resource unavailability, or technical issues. Daily standups help identify blockers for quick resolution."
    ),
    (
        "Define sprint goal.",
        ["Short statement describing sprint's primary objective",
         "Aligns team efforts for the sprint",
         "Helps with prioritisation decisions mid-sprint",
         "Provides context for daily work",
         "Set collaboratively during Sprint Planning"],
        "A sprint goal is a short statement describing the primary objective and value the team aims to deliver in the sprint. It aligns team efforts and provides context for daily work and prioritisation decisions."
    ),
    (
        "What is the role of Scrum Master in conflict resolution?",
        ["Facilitate discussion and help team understand perspectives",
         "Coach on resolution approaches",
         "Does not decide outcomes for the team",
         "Creates safe space for honest conversation",
         "Supports team in reaching agreement"],
        "When conflicts arise, the Scrum Master facilitates discussion, helps team members understand different perspectives, coaches them on resolution approaches, and supports agreement. They enable team resolution rather than deciding outcomes."
    ),
    (
        "Explain agile metrics.",
        ["Quantitative measures of agile project performance",
         "Include velocity, burndown, cycle time, lead time",
         "Reveal performance trends and improvement areas",
         "Support data-driven decision-making",
         "Should reflect team health, not just output"],
        "Agile metrics include velocity, burndown/burnup charts, cycle time, lead time, defect rate, and team satisfaction. They provide visibility into performance and trends, supporting data-driven decision-making."
    ),
    (
        "What is cycle time?",
        ["Duration from work start to completion",
         "Measures efficiency of the development process",
         "Shorter cycle time = faster value delivery",
         "Key metric for process improvement",
         "Different from lead time"],
        "Cycle time is the time from when work starts to when it's complete and deployed. Shorter cycle time indicates better efficiency and faster value delivery. It's a key metric for process improvement."
    ),
    (
        "Define lead time.",
        ["Time from requirement request to delivery to customer",
         "Includes analysis, development, testing, and deployment",
         "Shorter lead time = better responsiveness",
         "Customer-facing metric",
         "Different from cycle time"],
        "Lead time is the total time from when a requirement is requested to when it's delivered to the customer. Shorter lead time reflects better responsiveness to customer needs."
    ),
    (
        "What is kanban?",
        ["Visual work management with continuous flow",
         "Board with columns representing workflow stages",
         "WIP limits control flow and reduce context switching",
         "Emphasises continuous delivery",
         "No fixed sprints — pull-based system"],
        "Kanban is a visual management method using a board with columns representing workflow stages. WIP limits control flow and Kanban emphasises continuous delivery and process optimisation."
    ),
    (
        "Explain Scrum of Scrums.",
        ["Technique for coordinating multiple Scrum teams",
         "Representatives from each team meet regularly",
         "Discusses dependencies, integration, and shared challenges",
         "Typically facilitated by Scrum Masters",
         "Scales Scrum across a programme"],
        "Scrum of Scrums is a technique for coordinating multiple Scrum teams working on the same product. Representatives from each team meet to discuss dependencies, integration points, and shared challenges."
    ),
    (
        "What is a scaling framework in agile?",
        ["Methods for managing multiple agile teams",
         "Examples: SAFe, LeSS, Nexus, Spotify model",
         "Defines roles, ceremonies, and coordination mechanisms",
         "Aligns multiple teams toward common goals",
         "Selected based on organisation size and context"],
        "Scaling frameworks (SAFe, LeSS, Nexus, Spotify model) provide structures for coordinating multiple Scrum teams, defining roles, ceremonies, and coordination mechanisms to align teams toward common goals."
    ),
    (
        "Define SAFe (Scaled Agile Framework).",
        ["Enterprise framework for scaling agile across organisation",
         "Levels: Team, Program, Large Solution, Portfolio",
         "Integrates lean principles and DevOps",
         "Most widely adopted scaling framework",
         "Adds roles like RTE (Release Train Engineer)"],
        "SAFe is a comprehensive framework for scaling agile across enterprises with multiple levels: Team, Program, Large Solution, and Portfolio. It integrates lean principles, system thinking, and DevOps practices."
    ),
    (
        "What is LeSS (Large-Scale Scrum)?",
        ["Minimalist approach to scaling Scrum",
         "Applies Scrum principles at scale with minimal additions",
         "Avoids bureaucracy and extra roles",
         "Handles 2–8 teams typically",
         "Emphasises organisational simplicity"],
        "LeSS is a framework that applies Scrum principles at scale with minimal additional roles and artifacts. It emphasises agility while avoiding bureaucracy and typically handles 2–8 teams."
    ),
    (
        "Explain dependencies in Scrum.",
        ["Work reliance between teams or sprints",
         "Scrum Masters identify and manage dependencies",
         "Dependencies can block sprint delivery",
         "Visualised on dependency boards or risk maps",
         "Minimising dependencies improves flow"],
        "Dependencies in Scrum are connections between work items where one depends on another being completed. Scrum Masters identify and manage dependencies to minimise blocking, coordinate across teams, and maintain flow."
    ),
    (
        "What is technical excellence?",
        ["High-quality code and sustainable development practices",
         "Includes clean code, testing, CI, refactoring",
         "Sustains long-term team productivity",
         "Scrum Master fosters culture valuing technical quality",
         "Prevents accumulation of technical debt"],
        "Technical excellence involves clean code, comprehensive testing, continuous integration, refactoring, and architectural discipline. Scrum Masters should foster a culture that values technical excellence alongside delivering features."
    ),
    (
        "Define product vision.",
        ["Long-term goal and direction for the product",
         "Clear and compelling statement of purpose",
         "Describes target users and value proposition",
         "Product Owner owns the vision",
         "Scrum Master helps team align decisions with vision"],
        "Product vision is a clear, compelling statement describing the product's long-term goals, target users, and value proposition. The Scrum Master helps the team understand and align with the product vision."
    ),
    (
        "What is a stakeholder feedback loop?",
        ["Continuous communication of customer needs and preferences",
         "Sprint Reviews are primary feedback mechanism",
         "Customer interviews supplement Sprint Reviews",
         "Feedback incorporated into Product Backlog",
         "Ensures product evolves with user needs"],
        "Stakeholder feedback loop involves regularly collecting feedback from users and customers, sharing it with the team, and incorporating it into product development. Sprint Reviews facilitate this feedback loop."
    ),
    (
        "Explain minimum viable product (MVP).",
        ["Smallest product version with core value to customers",
         "Enables faster learning with reduced investment",
         "Gathers early adopter feedback",
         "Reduces waste on unnecessary features",
         "Foundation for iterative product development"],
        "MVP is the smallest version of the product with sufficient features to satisfy early adopters and gather user feedback. Building MVPs first enables faster learning and reduces waste on unnecessary features."
    ),
    (
        "What is iterative development?",
        ["Building product through repeated cycles of feedback",
         "Each iteration produces a working increment",
         "Feedback from each cycle informs the next",
         "Scrum sprints embody iterative development",
         "Enables continuous product refinement"],
        "Iterative development involves building the product in small increments, getting feedback, making improvements, and repeating. Scrum's sprint cycles embody iterative development, enabling continuous refinement."
    ),
    (
        "Define organisational impediments.",
        ["Obstacles outside team's direct control",
         "Include policy issues, resource allocation, external dependencies",
         "Scrum Master escalates to management",
         "Requires organisational support to resolve",
         "Common cause of team frustration"],
        "Organisational impediments are blockers outside the Scrum team's authority, such as policy issues or resource allocation. Scrum Masters escalate and facilitate resolution with management."
    ),
    (
        "What is the Scrum Master's role with management?",
        ["Coach management on agile principles",
         "Protect team autonomy from interference",
         "Escalate impediments requiring management action",
         "Manage stakeholder expectations",
         "Help organisation adapt to agile ways of working"],
        "Scrum Masters coach management on agile principles, explain why certain practices matter, protect team autonomy, escalate impediments, manage stakeholder expectations, and help the organisation adopt agile ways of working."
    ),
    (
        "Explain psychological safety in teams.",
        ["Environment where people feel safe taking risks",
         "Team members can speak up without fear of punishment",
         "Foundation for high-performing teams",
         "Scrum Master models vulnerability and encourages participation",
         "Enables honest retrospectives and innovation"],
        "Psychological safety is the team's belief that they can take risks, speak up, and admit mistakes without fear. Scrum Masters foster this by modelling vulnerability, encouraging participation, and protecting team members."
    ),
    (
        "What is intrinsic motivation?",
        ["Internal drive to do meaningful work",
         "Comes from autonomy, mastery, and purpose",
         "More sustainable than extrinsic motivation",
         "Scrum Masters increase it by ensuring teams understand value",
         "Supported by team autonomy and growth opportunities"],
        "Intrinsic motivation comes from finding work meaningful, having autonomy, mastering skills, and seeing purpose. Scrum Masters increase intrinsic motivation by ensuring teams understand value, have autonomy, and receive growth opportunities."
    ),
    (
        "Define trust-based relationships.",
        ["Foundation for high-performing teams",
         "Built through consistency, reliability, and transparency",
         "Enables better collaboration and faster decisions",
         "Scrum Master models trustworthy behaviour",
         "Takes time to build but is fragile"],
        "Trust-based relationships among team members and with leadership enable better collaboration, faster decision-making, and psychological safety. Scrum Masters build trust through consistency, reliability, and transparency."
    ),
    (
        "What is a retrospective anti-pattern?",
        ["Common dysfunction making retrospectives ineffective",
         "Examples: blame culture, no action items, same issues recurring",
         "Low participation or poor facilitation",
         "Scrum Master must recognise and correct these",
         "Anti-patterns prevent continuous improvement"],
        "Anti-patterns include blame-focused discussions, lack of action items, same issues recurring, limited participation, or poor facilitation. Scrum Masters should recognise and correct these to make retrospectives effective."
    ),
    (
        "Explain the concept of done in Agile.",
        ["Shared, explicit agreement on what finished means",
         "Without clear DoD teams accumulate incomplete work",
         "Applied rigorously to every user story",
         "Collaboratively defined by the team",
         "Reviewed and updated as team matures"],
        "Done in Agile is a shared, explicit Definition of Done of what it means for work to be finished. Scrum Masters ensure teams collaboratively define and apply DoD rigorously to prevent incomplete work accumulating."
    ),
    (
        "What is a story breakdown?",
        ["Decomposing large user stories into smaller ones",
         "Smaller stories fit within a single sprint",
         "Easier to estimate, implement, and test",
         "Uses vertical slicing to retain business value",
         "Performed during backlog refinement"],
        "Story breakdown is dividing large, complex user stories into smaller, more manageable stories that fit within a sprint. Smaller stories are easier to estimate, implement, test, and demonstrate."
    ),
    (
        "Define acceptance testing.",
        ["Validation that work meets acceptance criteria",
         "Performed by Product Owner, QA, or users",
         "Confirms work delivers expected value",
         "Last check before story is marked Done",
         "Based on criteria defined in the user story"],
        "Acceptance testing verifies that completed work meets the acceptance criteria defined in the User Story. It's typically performed by the Product Owner, QA, or users to ensure the work delivers expected value."
    ),
    (
        "What is the difference between velocity and capacity?",
        ["Velocity: historical completion rate in story points",
         "Capacity: available team time for a sprint",
         "Velocity is measured after the sprint",
         "Capacity is calculated before the sprint",
         "Capacity helps estimate achievable velocity"],
        "Velocity is the actual rate at which the team completes work (story points per sprint), while capacity is the available time the team has for work in a sprint. Capacity helps estimate velocity for future sprints."
    ),
    (
        "Explain sprint planning anti-patterns.",
        ["Over-committing beyond team capacity",
         "Poor prioritisation — pulling in low-value items",
         "Excluding team input in selection",
         "Inadequate story breakdown before the sprint",
         "No sprint goal set"],
        "Sprint planning anti-patterns include over-committing, poor prioritisation, excluding team input, and inadequate story breakdown. Scrum Masters should guide teams toward effective, realistic sprint planning."
    ),
    (
        "What is a standup anti-pattern?",
        ["Standup becomes status report to Scrum Master",
         "Exceeds 15-minute timebox",
         "Off-topic problem-solving discussions",
         "Low or inconsistent participation",
         "No follow-up on blockers raised"],
        "Standup anti-patterns include standups becoming status reports, excessive length, off-topic discussions, or low participation. Effective standups are peer-to-peer, focused on blockers, and strictly time-boxed."
    ),
    (
        "Define review-driven development.",
        ["Feedback from Sprint Review drives development direction",
         "Stakeholder feedback directly influences next priorities",
         "Ensures development aligns with user expectations",
         "Promotes customer-centric development",
         "Reduces risk of building unwanted features"],
        "Review-driven development means sprint review feedback from stakeholders directly influences upcoming priorities and development direction, ensuring development aligns with user needs and expectations."
    ),
    (
        "What is the Scrum Master's role in team dynamics?",
        ["Observe communication patterns and behaviours",
         "Coach on collaboration and effective working",
         "Address dysfunction early before it escalates",
         "Facilitate conflict resolution",
         "Help build effective working relationships"],
        "Scrum Masters watch for communication patterns, personality conflicts, knowledge silos, or disengagement. They coach teams on collaboration, facilitate conflict resolution, and help teams develop effective working relationships."
    ),
    (
        "Explain the concept of slack in development.",
        ["Intentionally reserved capacity for unexpected work",
         "Typically 10–20% of sprint capacity",
         "Used for refactoring, learning, urgent items",
         "Prevents overcommitment and burnout",
         "Scrum Master helps team understand value of slack"],
        "Slack is intentionally reserved capacity for unexpected issues, refactoring, or urgent work. Without slack, teams become overcommitted and stressed. Scrum Masters help teams understand and maintain healthy slack."
    ),
    (
        "What is a product roadmap in Scrum context?",
        ["High-level flexible plan of upcoming features and releases",
         "Spans several months or quarters",
         "Provides direction while remaining adaptable",
         "Owned by the Product Owner",
         "Reviewed and updated regularly"],
        "A product roadmap in Scrum is a flexible, high-level view of planned features, improvements, and releases over several months. It provides direction while remaining adaptable to changing market conditions and feedback."
    ),
    (
        "Define sprint commitment.",
        ["Team's explicit agreement on what they'll complete",
         "Based on historical velocity and capacity",
         "Increases team ownership and accountability",
         "Not a promise — a best-effort forecast",
         "Sprint goal defines what success looks like"],
        "Sprint commitment is the team's explicit agreement about what they'll complete in the sprint. It's based on historical velocity and honest assessment of capacity, increasing team ownership and accountability."
    ),
    (
        "What is the Scrum Master's role in quality?",
        ["Coach team on quality practices",
         "Ensures testing, code review, and CI/CD are adopted",
         "Removes obstacles to quality",
         "Discusses quality in retrospectives",
         "Helps team balance speed with quality"],
        "Scrum Masters ensure the team adopts quality practices (testing, code review, CI/CD), discusses quality in retrospectives, removes obstacles to quality, and helps the team balance speed with quality."
    ),
    (
        "Explain refactoring in Agile context.",
        ["Improving code structure without changing behaviour",
         "Reduces technical debt",
         "Maintains and improves code maintainability",
         "Scrum Master ensures time is allocated for it",
         "Part of sustainable development practices"],
        "Refactoring is improving code quality, readability, and maintainability without changing functionality. Scrum Masters should ensure teams dedicate time to refactoring to manage technical debt and maintain productivity."
    ),
    (
        "What is the concept of fail fast?",
        ["Quick learning from experiments and mistakes",
         "Accepts some failure as part of innovation",
         "Iterates rapidly based on results",
         "Requires psychological safety to work",
         "Reduces cost of failure through early detection"],
        "Fail fast means running quick experiments, accepting some failures as learning opportunities, and iterating rapidly based on results. Scrum Masters foster psychological safety to enable this mindset."
    ),
    (
        "Define risk management in Scrum.",
        ["Identifying and mitigating threats to sprint and product success",
         "Risks surface in retrospectives and daily standups",
         "Product Owner prioritises high-risk items early",
         "Scrum Master facilitates risk discussions",
         "Regular inspection reduces risk through early detection"],
        "Risk management in Scrum involves identifying potential threats, assessing probability and impact, and developing mitigation strategies. Regular retrospectives help identify and address emerging risks."
    ),
    (
        "What is a time-boxed activity?",
        ["Activity with a strict, respected time limit",
         "All Scrum ceremonies are time-boxed",
         "Improves focus and prevents excessive discussion",
         "Respects team members' time",
         "Scrum Master enforces timeboxes"],
        "Timeboxing means limiting an activity to a specific duration. All Scrum ceremonies are time-boxed. Timeboxing improves focus, prevents excessive discussion, and respects team time."
    ),
    (
        "Explain the role of documentation in Scrum.",
        ["Just-enough documentation for clarity and knowledge preservation",
         "Working software over comprehensive documentation (Agile Manifesto)",
         "Teams decide what documentation is essential",
         "Scrum Master helps team find the right balance",
         "Avoids both over-documentation and under-documentation"],
        "Scrum emphasises working software over comprehensive documentation, but teams still need essential documentation. Scrum Masters help teams find the right balance between documentation and delivery speed."
    ),
    (
        "What is the concept of emergence in Scrum?",
        ["Design and solutions emerging through iterations",
         "Not fully designed upfront",
         "Architecture evolves based on feedback and learning",
         "Reduces waste from over-engineering",
         "Enabled by iterative development cycles"],
        "Emergence means the product design and solutions emerge gradually through iterative development and feedback rather than being fully designed upfront. This adaptive approach reduces waste and enables innovation."
    ),
    (
        "Define value delivery in Scrum.",
        ["Regular, tangible business benefits to customers",
         "Measured by working increments, not activity",
         "Sprint goal defines the value to be delivered",
         "Product Owner prioritises highest value items",
         "Scrum Master ensures sprints focus on value"],
        "Value delivery in Scrum means teams regularly produce working increments that provide customer value. Scrum Masters ensure sprints focus on value delivery and measure progress by value, not just activity."
    ),
    (
        "What is the Scrum Master's role in organisational change?",
        ["Advocate for agile transformation",
         "Help organisation adopt agile mindset and culture",
         "Coach managers on supporting agile teams",
         "Drive continuous organisational improvement",
         "Champion agile practices beyond the team"],
        "Scrum Masters champion agile practices, help the organisation adopt agile mindset, coach managers on supporting agile teams, and drive continuous organisational improvement toward agile excellence."
    ),
    (
        "Explain the concept of Definition of Ready.",
        ["Criteria backlog items must meet before entering sprint",
         "Clear description and acceptance criteria",
         "Story must be estimated by the team",
         "No blocking external dependencies",
         "Prevents unprepared stories entering sprints"],
        "DoR establishes requirements for backlog items before they're committed to a sprint: clear description, acceptance criteria defined, estimated, and no blocking dependencies. DoR prevents stories being pulled into sprints unprepared."
    ),
    (
        "What is a process improvement in Scrum?",
        ["Changes to team working methods identified in retrospective",
         "Small, incremental changes over time",
         "Committed as action items in retrospective",
         "Tracked and reviewed each sprint",
         "Builds compound improvement over time"],
        "Process improvements are changes to team practices, workflows, or tools identified in retrospectives to increase efficiency, quality, or satisfaction. Scrum Masters facilitate identification and implementation of improvements."
    ),
    (
        "Define behavioural-driven development (BDD).",
        ["Collaborative approach linking business and technical language",
         "Uses human-readable specifications (Gherkin syntax)",
         "Describes system behaviour from user perspective",
         "Reduces misunderstandings between business and development",
         "Extends TDD with business context"],
        "BDD extends TDD by using human-readable specifications that describe behaviour. It improves collaboration between business stakeholders and developers, reducing misunderstandings and aligning development with business needs."
    ),
    (
        "What is the role of Scrum Master in onboarding?",
        ["Welcome new team members and explain Scrum practices",
         "Explain team norms, tools, and ways of working",
         "Pair new members with experienced colleagues",
         "Ensure new members understand the product and context",
         "Help them contribute effectively as quickly as possible"],
        "Scrum Masters welcome new team members, explain Scrum practices and team norms, pair them with experienced members, ensure they understand the product and context, and help them contribute effectively."
    ),
    (
        "Explain velocity forecasting.",
        ["Predicting future sprint completion using historical data",
         "Uses average or range of recent velocities",
         "Enables more accurate release planning",
         "Accounts for capacity variations",
         "Improves as team stabilises"],
        "Velocity forecasting uses historical velocity data to predict how many story points the team will likely complete in future sprints, enabling better release planning and realistic commitment."
    ),
    (
        "What is the Scrum Master's role in mentoring?",
        ["Coach individuals on agile practices",
         "Encourage skill development and career growth",
         "Share knowledge and experience",
         "Help individuals become more effective contributors",
         "Support team members' professional development"],
        "Scrum Masters mentor team members on agile practices, encourage skill development, support career growth, and help individuals become more effective contributors to the team."
    ),
    (
        "Define sprint variance.",
        ["Difference between planned and actual story points completed",
         "Small, consistent variance indicates reliable planning",
         "Large variance suggests estimation or planning issues",
         "Reviewed in retrospective for improvement",
         "Tracked over multiple sprints for trend analysis"],
        "Sprint variance is the difference between estimated and actual story points completed. Small, consistent variance indicates reliable planning. Large variance suggests estimation or planning issues to address in retrospective."
    ),
    (
        "What is the concept of Kaizen?",
        ["Continuous, incremental improvement philosophy",
         "Originates from Japanese manufacturing",
         "Embodied in Scrum through retrospectives",
         "Small improvements compound into significant gains",
         "Everyone contributes to improvement"],
        "Kaizen is a philosophy of continuous, incremental improvement. In Scrum, Kaizen is embodied through retrospectives where teams consistently identify and implement small improvements for significant compound benefits."
    ),
    (
        "Explain the role of Scrum Master in stakeholder management.",
        ["Communicate progress through reviews and reports",
         "Manage expectations about capabilities and timelines",
         "Facilitate stakeholder participation in key ceremonies",
         "Address stakeholder concerns proactively",
         "Bridge between team and business stakeholders"],
        "Scrum Masters keep stakeholders informed of progress through reviews and reports, manage expectations, facilitate stakeholder participation in ceremonies, and address stakeholder concerns proactively."
    ),
    (
        "What is the concept of servant leader?",
        ["Leadership style focused on serving and enabling others",
         "Prioritises others' growth and success",
         "Listens actively and removes obstacles",
         "Fosters trust and team autonomy",
         "Contrasts with traditional hierarchical leadership"],
        "Servant leadership prioritises others' growth and success over personal advancement. Servant leader Scrum Masters listen actively, support team development, remove obstacles, foster trust, and enable team autonomy."
    ),
    (
        "Define agile coaching.",
        ["Mentoring and facilitating organisational agile transformation",
         "Helps teams adopt and improve agile practices",
         "Guides teams through challenges and change",
         "Builds long-term agile capability",
         "Scrum Masters should develop coaching skills"],
        "Agile coaching involves helping teams adopt agile practices, guiding them through challenges, improving their processes, and building capability. Scrum Masters should develop coaching skills to maximise their impact."
    ),
    (
        "What is the Scrum Master's role in DevOps practices?",
        ["Promote CI/CD and automation practices",
         "Work with operations to remove deployment barriers",
         "Advocate for infrastructure as code",
         "Ensure rapid, reliable releases",
         "Bridge development and operations collaboration"],
        "Scrum Masters advocate for DevOps practices like CI/CD, infrastructure as code, and automated testing. They work with operations teams to remove deployment barriers and ensure rapid, reliable releases."
    ),
    (
        "Explain the concept of Whole Team Approach.",
        ["Everyone shares responsibility for quality and delivery",
         "No silos — developers, testers, designers all collaborate",
         "Everyone contributes to sprint goals",
         "Collective ownership of the product",
         "Scrum Master facilitates unified ownership"],
        "Whole Team Approach means all team members share responsibility for quality, delivery, and success. There are no silos; everyone collaborates toward sprint goals. Scrum Masters facilitate this unified ownership."
    ),
    (
        "What is the strategic importance of Scrum Master in organisation?",
        ["Drives organisational agility and team performance",
         "Ensures effective Scrum implementation",
         "Builds high-performing, self-organising teams",
         "Removes organisational impediments to flow",
         "Enables organisation to deliver value faster"],
        "The Scrum Master is critical to organisational success by ensuring effective Scrum implementation, building high-performing teams, removing organisational impediments, fostering continuous improvement culture, and enabling the organisation to deliver value faster."
    ),
]


# ── 100 Project Manager Q&A ───────────────────────────────────────────────
PROJECT_MANAGER_QA = [
    (
        "What is project management?",
        ["Planning, executing, and controlling work to achieve objectives",
         "Applies knowledge, skills, tools, and techniques to projects",
         "Involves managing scope, schedule, resources, and budget",
         "Monitors progress and controls changes",
         "Delivers successful outcomes meeting requirements"],
        "Project management is the application of knowledge, skills, tools, and techniques to project activities to meet project requirements. It involves planning scope, schedule, resources, and budget; executing work; monitoring progress; and controlling changes to deliver successful outcomes."
    ),
    (
        "Define a project.",
        ["Temporary endeavour to create a unique product or service",
         "Has defined beginning and end dates",
         "Specific objectives, budget, and resources",
         "Unlike operations — projects are unique and time-bound",
         "Results in deliverables (software, building, product launch)"],
        "A project is a temporary undertaking with defined beginning and end dates, specific objectives, budget, and resources. Unlike operations (ongoing work), projects are unique, time-bound, and result in deliverables such as building a bridge, implementing software, or launching a product."
    ),
    (
        "What are the main knowledge areas in project management?",
        ["10 knowledge areas defined in PMBOK Guide",
         "Integration, Scope, Schedule, Cost, Quality",
         "Resource, Communications, Risk, Procurement",
         "Stakeholder Management",
         "Each includes planning, executing, and monitoring processes"],
        "The PMBOK Guide defines 10 knowledge areas: Project Integration Management, Scope, Schedule, Cost, Quality, Resource, Communications, Risk, Procurement, and Stakeholder Management. Each includes planning, executing, and monitoring processes."
    ),
    (
        "Explain the project lifecycle.",
        ["Phases from initiation to closure",
         "Initiation: define project and gain authorization",
         "Planning: develop detailed plans",
         "Execution: perform the work",
         "Monitoring & Controlling, then Closing"],
        "The project lifecycle includes phases: Initiation (define project), Planning (develop plans), Execution (perform work), Monitoring & Controlling (track and adjust), and Closing (finalize and document). Each phase has specific activities and deliverables."
    ),
    (
        "What is a project charter?",
        ["Formally authorizes the project",
         "Issued by executive sponsor or senior management",
         "Defines high-level objectives and business case",
         "Grants project manager authority",
         "Starting point for all detailed planning"],
        "A project charter is a formal document that officially authorizes project existence, provides project manager authority, describes high-level objectives, business case, and success criteria. It is the starting point for planning."
    ),
    (
        "Define project scope.",
        ["Defines what is and isn't included in the project",
         "Specifies deliverables and success criteria",
         "Scope creep is a major project risk",
         "Clear scope prevents stakeholder misalignment",
         "Managed through formal change control"],
        "Project scope defines the boundaries of work: what will be delivered, what won't be included, deliverables, and success criteria. Scope creep (unauthorised additions) is a major project risk. Clear scope prevents misalignment."
    ),
    (
        "What is a Work Breakdown Structure (WBS)?",
        ["Hierarchical decomposition of project deliverables",
         "Organises scope into manageable work packages",
         "Shows relationships and dependencies",
         "Foundation for scheduling and budgeting",
         "Every deliverable should appear in the WBS"],
        "A WBS is a hierarchical breakdown of project deliverables and work packages. It organises project scope into manageable components, showing relationships and dependencies, and is the foundation for scheduling, budgeting, and resource allocation."
    ),
    (
        "Explain project schedule.",
        ["Timeline showing when activities occur and their dependencies",
         "Sequences activities and defines durations",
         "Identifies critical path and slack",
         "Shows interdependencies between tasks",
         "Baseline for schedule performance measurement"],
        "A project schedule sequences activities, defines durations, identifies dependencies, and determines the timeline for project completion. It shows critical path (longest sequence of dependent activities) and slack (flexibility) for activities."
    ),
    (
        "What is critical path?",
        ["Longest sequence of dependent activities",
         "Determines minimum project duration",
         "Activities on critical path have zero slack",
         "Delays on critical path delay the entire project",
         "Focus management attention here"],
        "The critical path is the longest series of dependent project activities that determines minimum project duration. Activities on critical path have zero slack; delays here delay the entire project. Identifying critical path helps focus management attention."
    ),
    (
        "Define project budget.",
        ["Estimated costs for project completion",
         "Includes labour, materials, equipment, and overhead",
         "Budget baseline used for performance measurement",
         "Tracking actual vs budgeted spend ensures control",
         "Prevents cost overruns through monitoring"],
        "A project budget is the monetary estimate for project expenses including labour, materials, equipment, and overhead. Budget baseline is the approved budget; tracking actual vs budgeted spend ensures cost control and prevents overruns."
    ),
    (
        "What is project quality?",
        ["Meeting specifications and satisfying stakeholder needs",
         "Deliverables must meet acceptance criteria",
         "Quality management: planning, assurance, and control",
         "Prevention is cheaper than defect correction",
         "Quality is everyone's responsibility"],
        "Project quality means delivering deliverables that meet stated specifications, acceptance criteria, and stakeholder expectations. Quality management includes planning (defining standards), assurance (prevention), and control (detection) activities."
    ),
    (
        "Explain risk management.",
        ["Identifying, analysing, and responding to uncertainties",
         "Threats (negative) and opportunities (positive)",
         "Responses: avoid, mitigate, transfer, accept",
         "Risk monitored throughout project lifecycle",
         "Early identification prevents crisis management"],
        "Risk management involves identifying potential threats and opportunities, assessing probability and impact, planning responses (avoid, mitigate, transfer, accept), and monitoring risks throughout the project lifecycle. Early risk identification prevents crisis."
    ),
    (
        "What is a risk register?",
        ["Repository of identified risks and response plans",
         "Includes description, probability, impact, and owner",
         "Maintained throughout the project",
         "Provides visibility on risk status",
         "Enables effective, proactive risk management"],
        "A risk register is a repository of identified risks including description, probability, impact, owner, and response strategy. It is maintained throughout the project, providing visibility on risk status and enabling effective risk management."
    ),
    (
        "Define stakeholder management.",
        ["Identifying and engaging all interested parties",
         "Analyse interest level and influence of each stakeholder",
         "Plan tailored engagement strategies",
         "Maintain regular communication",
         "Ensure support and manage concerns proactively"],
        "Stakeholder management involves identifying all stakeholders, analysing their interests and influence, planning engagement strategies, and maintaining communication to ensure support and manage concerns throughout the project."
    ),
    (
        "What is a RACI matrix?",
        ["Maps roles and responsibilities for project activities",
         "R = Responsible (does the work)",
         "A = Accountable (final authority)",
         "C = Consulted (provides input)",
         "I = Informed (kept updated)"],
        "A RACI matrix defines roles for project activities: Responsible (does the work), Accountable (final authority), Consulted (provides input), Informed (kept updated). It prevents confusion about accountability and ensures clarity."
    ),
    (
        "Explain project communications.",
        ["Exchange of information among all project participants",
         "Determine stakeholder information needs",
         "Plan communication frequency and method",
         "Monitor communication effectiveness",
         "Manage stakeholder engagement throughout"],
        "Project communications involves determining information needs, planning how information flows, executing communication, monitoring that communication is effective, and managing stakeholder engagement throughout the project."
    ),
    (
        "What is a status report?",
        ["Periodic snapshot of project health",
         "Shows progress against plan and key metrics",
         "Includes risks, issues, completed work, next steps",
         "Keeps stakeholders informed",
         "Enables early intervention when problems arise"],
        "A status report provides a snapshot of project health: progress against plan, metrics (schedule variance, cost variance), risks and issues, completed work, and next steps. Status reports keep stakeholders informed and enable early intervention."
    ),
    (
        "Define earned value management (EVM).",
        ["Integrates scope, schedule, and cost performance",
         "Compares Planned Value, Earned Value, Actual Cost",
         "SV = EV - PV; CV = EV - AC",
         "SPI = EV/PV; CPI = EV/AC",
         "Enables objective performance measurement"],
        "EVM measures project performance by comparing planned value (PV), earned value (EV), and actual cost (AC). Key metrics include Schedule Variance (EV-PV), Cost Variance (EV-AC), Schedule Performance Index (EV/PV), and Cost Performance Index (EV/AC)."
    ),
    (
        "What is variance analysis?",
        ["Comparing actual vs planned performance",
         "Examines schedule, cost, and scope variance",
         "Identifies root causes of deviations",
         "Determines corrective actions needed",
         "Enables proactive project management"],
        "Variance analysis examines differences between planned and actual results (schedule variance, cost variance, scope variance) to understand causes and determine corrective actions. Regular variance analysis enables proactive management."
    ),
    (
        "Explain change control process.",
        ["Formal procedure for managing project changes",
         "Receive, log, and analyse change requests",
         "Assess impact on scope, schedule, cost, and risk",
         "Obtain formal approval before implementing",
         "Communicate approved changes to all stakeholders"],
        "Change control involves receiving change requests, analysing impact (scope, schedule, cost, risk), obtaining approvals, implementing approved changes, and communicating changes to stakeholders. Controls prevent scope creep and ensure change traceability."
    ),
    (
        "What is scope creep?",
        ["Unauthorised expansion of project boundaries",
         "Consumes resources, delays schedule, increases costs",
         "Caused by unclear requirements or weak change control",
         "Prevented through strict change control process",
         "Clear requirements and stakeholder management help"],
        "Scope creep is uncontrolled expansion of project scope beyond what was approved. It consumes resources, delays schedule, and increases costs. Preventing scope creep requires clear requirements, strict change control, and stakeholder management."
    ),
    (
        "Define baseline in project management.",
        ["Approved version of plan used for performance measurement",
         "Three baselines: scope, schedule, and cost",
         "Changes require formal change control",
         "Maintained to preserve performance measurement integrity",
         "Deviations from baseline trigger corrective action"],
        "Project baselines (scope, schedule, cost) are approved versions of plans used to measure project performance. Baselines should not change frequently; changes require formal change control to maintain integrity."
    ),
    (
        "What is a milestone?",
        ["Significant event or completion point in project",
         "Has zero duration — it is a point in time",
         "Helps track progress at key stages",
         "Often aligns with phase gates or delivery dates",
         "Used to communicate progress to stakeholders"],
        "Milestones mark important project events or completion points (e.g., requirements approval, design completion, testing start). Milestones have no duration, help track progress, and often align with phase gates or delivery dates."
    ),
    (
        "Explain dependencies in project context.",
        ["Relationships between project activities",
         "Finish-to-Start: most common — B starts after A finishes",
         "Start-to-Start: B starts when A starts",
         "Finish-to-Finish: B finishes when A finishes",
         "Identifying dependencies produces realistic schedules"],
        "Dependencies show how activities relate: Finish-to-Start (activity B starts after A finishes), Start-to-Start (B starts when A starts), Finish-to-Finish (B finishes when A finishes). Identifying dependencies helps develop realistic schedules."
    ),
    (
        "What is resource levelling?",
        ["Smoothing resource usage to resolve conflicts",
         "Adjusts schedule when resources are overallocated",
         "May extend project schedule as a result",
         "Produces achievable plans based on availability",
         "Prevents team burnout from overallocation"],
        "Resource levelling adjusts the schedule to resolve resource conflicts where the same person is assigned to overlapping tasks. While it may extend schedule, it produces realistic, achievable plans based on actual resource availability."
    ),
    (
        "Define procurement management.",
        ["Planning and acquiring external products or services",
         "Includes source selection and vendor evaluation",
         "Contracting: negotiating favourable terms",
         "Managing vendor performance throughout",
         "Good procurement saves money and reduces risk"],
        "Procurement management involves planning what to procure externally, source selection (identifying qualified vendors), contracting (negotiating terms), and managing vendor performance. Good procurement saves money and reduces risk."
    ),
    (
        "What is vendor management?",
        ["Supervising and maintaining vendor relationships",
         "Selecting qualified vendors through evaluation",
         "Negotiating favourable contract terms",
         "Monitoring performance against quality, schedule, cost",
         "Managing disputes and maintaining positive relationships"],
        "Vendor management includes selecting qualified vendors, negotiating favourable terms, monitoring vendor performance (quality, schedule, cost), managing disputes, and maintaining positive relationships for long-term benefit."
    ),
    (
        "Explain contract types.",
        ["Fixed-Price: cost fixed regardless of effort",
         "Cost-Plus: reimburse costs plus agreed fee",
         "Time and Materials: pay for time and materials used",
         "Each type allocates risk differently",
         "Contract type chosen based on project uncertainty"],
        "Common contract types: Fixed-Price (cost fixed regardless of effort), Cost-Plus (reimburse costs plus fee), Time and Materials (pay for time and materials used). Each has different risk allocation between project and vendor."
    ),
    (
        "What is quality assurance?",
        ["Systematic evaluation of work processes",
         "Preventive — focused on process improvement",
         "Different from Quality Control (detection)",
         "Includes audits and process reviews",
         "Ensures processes produce quality results"],
        "Quality Assurance involves implementing and monitoring processes to ensure they produce quality results. QA is preventive, focused on process improvement. It differs from Quality Control which detects defects after they occur."
    ),
    (
        "Define quality control.",
        ["Testing and inspecting deliverables against standards",
         "Detective — finds defects after they occur",
         "Verifies deliverables meet acceptance criteria",
         "Different from QA (which is preventive)",
         "Results feed lessons learned and process improvement"],
        "Quality Control involves inspecting, testing, and verifying deliverables to ensure they meet quality standards and acceptance criteria. QC is detective; it finds defects after they occur."
    ),
    (
        "What is the role of a project manager?",
        ["Plans, executes, monitors, and controls the project",
         "Organises resources and directs execution",
         "Monitors progress against approved plans",
         "Controls changes and scope",
         "Manages stakeholders and ensures quality"],
        "The Project Manager is responsible for overall project success: developing plans, organising resources, directing execution, monitoring progress against plans, controlling changes and scope, managing stakeholders, and ensuring quality."
    ),
    (
        "Explain project manager competencies.",
        ["Technical knowledge of the project domain",
         "Leadership — inspiring and motivating team",
         "Communication — clear, timely information exchange",
         "Problem-solving — addressing issues decisively",
         "Negotiation and business acumen"],
        "PM competencies include: technical knowledge, leadership, communication, problem-solving, negotiation, and business acumen. Competent PMs combine these to navigate complexity and deliver successful outcomes."
    ),
    (
        "What is stakeholder analysis?",
        ["Identifying and understanding all stakeholders",
         "Assessing interest level and influence of each",
         "Planning tailored engagement strategies",
         "High-influence, high-interest stakeholders need most attention",
         "Updated as project progresses and stakeholders change"],
        "Stakeholder analysis involves identifying who is affected by or affects the project, assessing their interest and influence, and planning engagement strategies. High-influence, high-interest stakeholders need the most attention."
    ),
    (
        "Define power/interest grid.",
        ["Tool for categorising stakeholders by power and interest",
         "High power, high interest: manage closely",
         "Low power, high interest: keep informed",
         "High power, low interest: keep satisfied",
         "Low power, low interest: monitor only"],
        "The power/interest grid plots stakeholders on axes of power (ability to influence) and interest (concern about project). This helps determine appropriate engagement strategies for each stakeholder group."
    ),
    (
        "What is integration management?",
        ["Ensures all project components work together cohesively",
         "Coordinates scope, schedule, cost, resources, quality",
         "PM is responsible for integration across all areas",
         "Includes developing project plan and directing execution",
         "Manages changes and controls integration"],
        "Integration management coordinates all project elements to create a cohesive project plan and ensure components work together. The PM is responsible for integration across all knowledge areas."
    ),
    (
        "Explain the concept of project constraints.",
        ["Limitations affecting project execution",
         "Scope, schedule, budget, and quality",
         "Constraints are often competing with each other",
         "PM must balance trade-offs between constraints",
         "Stakeholders must agree on priority of constraints"],
        "Project constraints include scope (what must be delivered), schedule (when it must be done), budget (available funding), and quality (standards to meet). Constraints are often competing; PM must balance trade-offs."
    ),
    (
        "What is the triple constraint?",
        ["Interdependency of scope, time, and cost",
         "Also called the iron triangle",
         "Increasing scope requires more time or budget",
         "Shortening schedule may reduce scope or increase cost",
         "Changing one constraint affects the others"],
        "The triple constraint shows that scope, schedule, and cost are interdependent: increasing scope requires more time/budget, shortening schedule may reduce scope or increase cost. Changing one always affects the others."
    ),
    (
        "Define project governance.",
        ["Structure ensuring accountability and strategic alignment",
         "Defines authority and decision-making processes",
         "Establishes oversight mechanisms",
         "Ensures project aligns with organisational strategy",
         "Provides escalation path for issues"],
        "Project governance defines authority, decision-making processes, oversight mechanisms, and accountability. It ensures the project aligns with organisational strategy, maintains appropriate controls, and escalates issues appropriately."
    ),
    (
        "What is a lessons learned session?",
        ["Reflection on project outcomes and processes",
         "What worked well, what didn't, what to improve",
         "Typically held at project closure",
         "Improves organisational knowledge and maturity",
         "Output shared for future project benefit"],
        "Lessons learned sessions, typically held at project closure, reflect on what worked well, what didn't, and improvements for future projects. They improve organisational knowledge and process maturity."
    ),
    (
        "Explain issue management.",
        ["Process for identifying and resolving problems",
         "Assess impact and assign ownership",
         "Develop and implement solutions",
         "Track issues until fully resolved",
         "Issue log maintains visibility on all problems"],
        "Issue management involves identifying problems affecting the project, assessing impact, assigning ownership, developing solutions, implementing resolutions, and tracking until closure. Issue logs maintain visibility on problems."
    ),
    (
        "What is a project sponsor?",
        ["Senior executive who champions the project",
         "Secures funding and resources",
         "Authorises the project charter",
         "Supports the PM and removes organisational obstacles",
         "Represents business case at executive level"],
        "A project sponsor is a senior executive who champions the project, secures funding and resources, authorises the project charter, supports the PM, and represents the business case. Sponsor involvement is critical to project success."
    ),
    (
        "Define stakeholder communication plan.",
        ["Strategy specifying how stakeholders are kept informed",
         "Defines what information, frequency, and method",
         "Identifies responsible parties for each communication",
         "Prevents misalignment and disengagement",
         "Updated as stakeholder needs change"],
        "A communication plan specifies what information stakeholders need, frequency of updates, communication methods, and responsible parties. Effective communication prevents misalignment and keeps stakeholders engaged."
    ),
    (
        "What is configuration management?",
        ["Tracking and controlling project deliverable versions",
         "Maintains records of what has been delivered",
         "Ensures team works from current versions",
         "Provides audit trail of all changes",
         "Prevents version control confusion"],
        "Configuration management maintains records of what has been delivered, changes made, and current versions. It ensures the team works from current versions and provides an audit trail of changes."
    ),
    (
        "Explain the concept of gold plating.",
        ["Adding unnecessary features beyond requirements",
         "Consumes resources and time without proportional value",
         "Not requested by stakeholders",
         "PM must prevent it by controlling scope",
         "Focus must remain on specified requirements"],
        "Gold plating is adding features not required by stakeholders, consuming resources and time without adding proportional value. PMs must prevent gold plating by controlling scope and keeping focus on specified requirements."
    ),
    (
        "What is risk response planning?",
        ["Developing strategies to address identified risks",
         "Avoid: eliminate the risk entirely",
         "Mitigate: reduce probability or impact",
         "Transfer: shift risk to a third party",
         "Accept: acknowledge and plan contingency"],
        "Risk response planning includes: Avoid (eliminate risk), Mitigate (reduce probability/impact), Transfer (shift risk to third party), Accept (acknowledge and plan contingency). Each risk has an assigned owner responsible for executing the response."
    ),
    (
        "Define contingency reserve.",
        ["Additional budget held for identified risks",
         "Distinct from management reserve (for scope changes)",
         "Amount depends on project uncertainty",
         "Released when risks materialise",
         "Controlled by project manager"],
        "Contingency reserve is additional budget held to address identified risks and unknown unknowns. It is distinct from management reserve. The amount depends on project uncertainty and organisational risk tolerance."
    ),
    (
        "What is project recovery?",
        ["Actions to bring a troubled project back on track",
         "Assess what went wrong and root causes",
         "Develop corrective and preventive actions",
         "Re-plan if necessary and realign resources",
         "Improve communication and restore stakeholder confidence"],
        "Project recovery involves assessing what went wrong, developing corrective and preventive actions, re-planning if necessary, realigning resources, improving communication, and implementing changes to restore project health."
    ),
    (
        "Explain project closure activities.",
        ["Finalising all deliverables and obtaining sign-off",
         "Transferring deliverables to operations",
         "Releasing project resources formally",
         "Archiving project documentation",
         "Conducting lessons learned session"],
        "Project closure includes: finalising deliverables, obtaining stakeholder acceptance and sign-off, transferring deliverables to operations, releasing resources, archiving project documentation, conducting lessons learned, and formally closing the project."
    ),
    (
        "What is handover/transition?",
        ["Transferring completed deliverables to operations",
         "Documenting support procedures and runbooks",
         "Training operational support staff",
         "Establishing ongoing support arrangements",
         "Ensuring business continuity post-project"],
        "Handover involves transferring completed deliverables to operational support, documenting procedures, training support staff, establishing support arrangements, and ensuring continuity. Poor handover causes post-project problems."
    ),
    (
        "Define project success criteria.",
        ["Measures that determine whether project is successful",
         "On-time delivery, within budget, meets quality standards",
         "Achieves intended business objectives",
         "Stakeholder satisfaction confirmed",
         "Success extends beyond schedule and budget to value realisation"],
        "Success criteria define how success will be measured: on-time delivery, within budget, meets quality standards, achieves business objectives, stakeholder satisfaction. Project success extends beyond schedule and budget to business value realisation."
    ),
    (
        "What is a project portfolio?",
        ["Collection of projects aligned with organisational strategy",
         "Portfolio management prioritises projects",
         "Balances resources across competing projects",
         "Aligns execution with strategic objectives",
         "Ensures the right projects are executed"],
        "A project portfolio is a collection of projects selected for execution to achieve organisational strategic objectives. Portfolio management prioritises projects, balances resources, and aligns execution with strategy."
    ),
    (
        "Explain programme management.",
        ["Managing related projects with shared objectives",
         "Coordinates multiple interdependent projects",
         "Manages cross-project dependencies",
         "Coordinates resource sharing",
         "Delivers greater value than individual projects alone"],
        "Programme management coordinates multiple related projects sharing common objectives or resources. Programme managers ensure projects align, manage cross-project dependencies, and coordinate resource sharing to deliver greater overall value."
    ),
    (
        "What is strategic alignment?",
        ["Ensuring projects directly support organisational goals",
         "Misaligned projects waste resources",
         "PM should understand organisational strategy",
         "Projects should be evaluated against strategic fit",
         "Regular reviews ensure continued alignment"],
        "Strategic alignment ensures projects and initiatives directly contribute to organisational strategy and objectives. Misaligned projects waste resources. PMs should understand strategy and keep projects aligned throughout."
    ),
    (
        "Define project selection criteria.",
        ["Factors used to choose which projects to execute",
         "Strategic fit — alignment with objectives",
         "Financial return — ROI and payback period",
         "Resource requirements and capacity",
         "Risk level and dependencies"],
        "Project selection considers: strategic fit, financial return (ROI, payback period), resource requirements, risk level, dependencies, urgency, and organisational capacity. Selection criteria ensure the portfolio supports organisational strategy."
    ),
    (
        "What is organisational project management (OPM)?",
        ["Strategic framework integrating projects with organisational strategy",
         "Integrates project, programme, and portfolio management",
         "Ensures projects contribute to business strategy",
         "Creates sustainable competitive advantage",
         "Aligns execution with organisational goals"],
        "OPM integrates project, programme, and portfolio management with organisational strategy and operations. It ensures projects contribute to business strategy and creates sustainable competitive advantage."
    ),
    (
        "Explain the role of PMO (Project Management Office).",
        ["Centralised function providing project support and oversight",
         "Standardises project management practices and tools",
         "Provides training and governance",
         "PMO types: supportive, controlling, or directive",
         "Strong PMO improves organisational project success rates"],
        "A PMO provides standardised project management practices, tools, training, and oversight. PMO roles vary: supportive (provides guidance), controlling (requires compliance), or directive (manages projects directly). Strong PMO improves project success rates."
    ),
    (
        "What is project management methodology?",
        ["Structured approach and framework for managing projects",
         "Examples: Waterfall, Agile, Hybrid",
         "Includes processes, templates, and tools",
         "Should fit project characteristics and culture",
         "Provides consistency across the organisation"],
        "A methodology provides framework, processes, templates, and tools for project management. Examples: Waterfall (sequential phases), Agile (iterative), Hybrid (combining both). Methodology should fit project characteristics and organisational culture."
    ),
    (
        "Define predictive project management.",
        ["Planning-focused approach with limited change tolerance",
         "Also called Waterfall",
         "Detailed requirements defined upfront",
         "Sequential phases with formal gates",
         "Best for stable requirements and regulatory environments"],
        "Predictive (Waterfall) emphasises upfront planning, detailed requirements, and limited changes. Best for stable requirements, regulatory environments, or projects where scope is well understood upfront."
    ),
    (
        "What is adaptive project management?",
        ["Change-tolerant approach with continuous learning",
         "Also called Agile",
         "Emphasises rapid iterations and flexibility",
         "Responds to change over following a fixed plan",
         "Best for uncertain requirements and evolving markets"],
        "Adaptive (Agile) emphasises rapid iterations, flexibility, and responding to change. Best for uncertain requirements, rapidly evolving markets, or products requiring frequent customer feedback."
    ),
    (
        "Explain hybrid project approaches.",
        ["Combines predictive and adaptive elements",
         "Structured planning with iterative delivery",
         "Leverages benefits of both approaches",
         "Provides planning certainty with flexibility to adapt",
         "Increasingly common in large organisations"],
        "Hybrid approaches combine structured planning (predictive) with iterative delivery (adaptive). This enables organisations to leverage planning certainty while maintaining the flexibility to adapt to change."
    ),
    (
        "What is PMBOK?",
        ["Project Management Body of Knowledge — published by PMI",
         "Comprehensive guide to project management terminology",
         "Defines processes, knowledge areas, and best practices",
         "Foundation for PMP certification",
         "Widely adopted as industry standard"],
        "PMBOK is a comprehensive guide published by PMI containing standard project management terminology, processes, knowledge areas, and best practices. It is the foundation for PMP certification and widely used in the industry."
    ),
    (
        "Define maturity models.",
        ["Frameworks assessing organisational project management capability",
         "Examples: CMM, OPM3",
         "Assess processes, people, tools, and culture",
         "Identify improvement areas and guide capability development",
         "Higher maturity correlates with better project outcomes"],
        "Maturity models (CMM, OPM3) assess organisational project management capability across dimensions such as processes, people, tools, and culture. They help identify improvement areas and guide capability development."
    ),
    (
        "What is a project health dashboard?",
        ["Visual representation of project status and key metrics",
         "Displays schedule, cost, quality, risks, and issues",
         "Uses status indicators: red/yellow/green (RAG)",
         "Enables quick assessment by stakeholders",
         "Identifies areas needing immediate attention"],
        "A health dashboard displays key project metrics (schedule, cost, quality, risks, issues) with status indicators (red/yellow/green). Dashboards enable quick assessment of project health and identify areas needing attention."
    ),
    (
        "Explain project forecasting.",
        ["Predicting project completion status using current data",
         "Uses Earned Value metrics for objective forecasting",
         "Forecast at Completion (FAC) estimated from CPI",
         "Helps stakeholders understand likely outcomes",
         "Enables early corrective action"],
        "Forecasting uses current data and trends to predict project outcomes (final cost, completion date). Forecast at Completion (FAC) estimated from Earned Value metrics helps stakeholders understand likely project results."
    ),
    (
        "What is a project audit?",
        ["Independent review of project compliance and performance",
         "Reviews adherence to plans, standards, and procedures",
         "Can be internal (PMO) or external (client or regulator)",
         "Provides assurance project is being managed properly",
         "Identifies improvement opportunities"],
        "An audit reviews project adherence to approved plans, standards, and procedures. Audits may be internal (PMO) or external (client). They provide assurance the project is being managed properly and identify improvement opportunities."
    ),
    (
        "Define request for proposal (RFP).",
        ["Document soliciting detailed vendor proposals",
         "Describes project work and requirements",
         "Specifies evaluation criteria for selection",
         "Attracts competitive bids",
         "Enables fair and documented vendor comparison"],
        "An RFP is a detailed document describing project work, requirements, evaluation criteria, and invites vendors to submit proposals. Well-written RFPs attract competitive bids and enable fair vendor comparison."
    ),
    (
        "What is bid evaluation?",
        ["Assessment of vendor proposals against criteria",
         "Compares cost, capability, quality, schedule, and experience",
         "Multiple evaluators provide objective assessment",
         "Process must be documented and transparent",
         "Leads to vendor selection and contract award"],
        "Bid evaluation compares proposals against criteria (cost, capability, quality, schedule, experience) to select the best vendor. Multiple evaluation team members provide objective assessment. The process should be documented and transparent."
    ),
    (
        "Explain project kickoff meeting.",
        ["Initial meeting launching project execution",
         "Establishes common understanding among all stakeholders",
         "Covers objectives, roles, schedule overview, success criteria",
         "Defines communication protocols",
         "Strong kickoff sets a positive and aligned tone"],
        "The kickoff meeting brings all project stakeholders together to establish common understanding: project objectives, roles and responsibilities, schedule overview, success criteria, and communication protocols. A strong kickoff sets a positive tone."
    ),
    (
        "What is a scope statement?",
        ["Detailed document defining what will and won't be delivered",
         "Describes deliverables, acceptance criteria, and exclusions",
         "Documents constraints and assumptions",
         "Basis for scheduling, budgeting, and resource planning",
         "Reviewed and approved by key stakeholders"],
        "The scope statement describes project deliverables, acceptance criteria, constraints, assumptions, and exclusions. It is the detailed version of high-level scope and provides the basis for scheduling, budgeting, and resource planning."
    ),
    (
        "Define product requirements document (PRD).",
        ["Detailed description of what the product should do",
         "Specifies features, functionality, and performance requirements",
         "Includes UI requirements, constraints, and acceptance criteria",
         "Input to project planning and team execution",
         "Shared with all team members and stakeholders"],
        "PRD specifies product features, functionality, performance requirements, user interface requirements, constraints, and acceptance criteria. PRD is input to project planning and guides team execution."
    ),
    (
        "What is a project organisation chart?",
        ["Shows roles, reporting relationships, and authority levels",
         "Clarifies who reports to whom",
         "Defines decision-making authority",
         "Shows communication channels within the team",
         "Prevents confusion about accountability"],
        "Project org chart shows roles, reporting relationships, and authority levels within the project. It clarifies who reports to whom, decision-making authority, and communication channels. Clear organisation prevents confusion."
    ),
    (
        "Explain project staffing plan.",
        ["Strategy for acquiring and organising project resources",
         "Identifies skills needed for each phase",
         "Covers internal vs external acquisition strategy",
         "Timeline for when resources are needed",
         "Resource allocation across project phases"],
        "The staffing plan identifies resource requirements, specific skills needed, acquisition strategy (internal vs external), timeline for resource acquisition, and resource allocation across project phases."
    ),
    (
        "What is resource allocation?",
        ["Assignment of specific resources to project activities",
         "Considers resource availability, skills, and constraints",
         "Maximises utilisation while preventing overallocation",
         "Visualised using resource histograms",
         "Updated as project progresses"],
        "Resource allocation assigns specific people, equipment, or budget to activities. Allocation considers resource availability, skills, and constraints. Good allocation maximises resource utilisation while preventing overallocation."
    ),
    (
        "Define training and development in projects.",
        ["Building team capabilities for project success",
         "Technical skills, tools, and methodology training",
         "Investment improves project execution",
         "Increases team capability and confidence",
         "Enhances team retention and satisfaction"],
        "Training and development strengthens team skills in technical areas, tools, or project methodologies. Investment in team development improves project execution, increases team capability, and enhances retention."
    ),
    (
        "What is team morale?",
        ["Overall attitude and satisfaction of the project team",
         "High morale increases productivity and quality",
         "Improved by recognition, growth, and clear communication",
         "PM creates positive work environment",
         "Conflict resolution and fairness are critical"],
        "High team morale increases productivity, quality, and retention. PMs improve morale by recognising achievements, providing growth opportunities, maintaining clear communication, resolving conflicts, and creating a positive work environment."
    ),
    (
        "Explain conflict resolution.",
        ["Managing disagreements among team members or stakeholders",
         "Collaborating: finding win-win solutions (best)",
         "Compromising: splitting differences",
         "Avoiding: deferring the issue",
         "Accommodating or Forcing depending on situation"],
        "Conflict resolution approaches include: Collaborating (finding win-win solutions), Compromising (splitting differences), Avoiding (deferring), Accommodating (yielding), and Forcing (imposing a decision). Best approach depends on the situation."
    ),
    (
        "What is motivation theory in PM context?",
        ["Frameworks for understanding what drives team performance",
         "Maslow: hierarchy of needs from basic to self-actualisation",
         "Herzberg: hygiene factors vs motivators",
         "McGregor: Theory X (control) vs Theory Y (trust)",
         "PMs apply these to motivate teams effectively"],
        "Motivation theories (Maslow, Herzberg, McGregor) help PMs understand what motivates team members (achievement, recognition, responsibility, growth). Motivated teams consistently perform better."
    ),
    (
        "Define leadership styles.",
        ["Different approaches to directing and influencing others",
         "Directive: telling people what to do",
         "Supportive: encouraging and coaching",
         "Participative: involving team in decisions",
         "Effective PMs adapt style to situation and person"],
        "Leadership styles include: Directive (telling), Supportive (encouraging), Participative (involving), Achievement-oriented (challenging). Effective PMs adapt their style to the situation and team members. Situational leadership improves effectiveness."
    ),
    (
        "What is emotional intelligence?",
        ["Ability to understand and manage emotions",
         "Self-awareness, self-regulation, motivation, empathy",
         "Social skills enabling effective relationships",
         "PMs with high EI are more effective leaders",
         "Improves conflict resolution and team communication"],
        "Emotional intelligence (EI) includes self-awareness, self-regulation, motivation, empathy, and social skills. High EI enables better interpersonal relationships, conflict resolution, and team leadership, making PMs more effective."
    ),
    (
        "Explain negotiation skills.",
        ["Ability to reach mutually beneficial agreements",
         "Used for resource securing, vendor management, conflict resolution",
         "Interest-based negotiation focuses on interests, not positions",
         "Preparation and understanding BATNA is critical",
         "Produces better outcomes than positional bargaining"],
        "Negotiation skills help PMs resolve conflicts, secure resources, manage vendor relationships, and negotiate with stakeholders. Interest-based negotiation (focusing on interests, not positions) usually produces better outcomes."
    ),
    (
        "What is crisis management?",
        ["Rapid response to unexpected emergencies threatening the project",
         "Assess situation quickly and accurately",
         "Activate crisis response plan",
         "Communicate clearly to all stakeholders",
         "Manage recovery and document lessons learned"],
        "Crisis management involves rapid response to unexpected, serious situations threatening the project (resource loss, technical failure, scope emergency). Effective crisis management includes: assess situation, activate response plan, communicate clearly, and manage recovery."
    ),
    (
        "Define project metrics.",
        ["Quantitative measures of project performance",
         "SPI, CPI, defect rates, productivity, ROI",
         "Enable data-driven decision-making",
         "Identify performance trends over time",
         "Reported regularly to stakeholders"],
        "Metrics track project health: Schedule Performance Index, Cost Performance Index, defect rates, productivity, quality, and ROI. Metrics enable data-driven decision-making and identify trends requiring intervention."
    ),
    (
        "What is trend analysis?",
        ["Examining data patterns over time to predict future performance",
         "Reviews historical project performance data",
         "Identifies patterns indicating problems developing",
         "Cost increasing or schedule slipping are common signals",
         "Enables proactive intervention before problems escalate"],
        "Trend analysis reviews historical project data to identify patterns and predict future performance. Trends indicating problems (rising costs, slipping schedule) require intervention before they become critical issues."
    ),
    (
        "Explain project documentation standards.",
        ["Guidelines for creating consistent, high-quality documents",
         "Specifies format, content, naming conventions",
         "Covers storage, retrieval, and version control",
         "Improves document quality and usability",
         "Ensures consistency across the project"],
        "Documentation standards ensure project documents are consistent, complete, and professional. Standards specify format, content requirements, naming conventions, storage, and version control. Consistent documentation improves quality."
    ),
    (
        "What is project knowledge management?",
        ["Capturing and sharing project learning and expertise",
         "Documenting lessons learned and best practices",
         "Sharing solutions to common problems",
         "Improves future project execution",
         "Builds organisational project management capability"],
        "Knowledge management involves documenting project knowledge, lessons learned, best practices, and solutions. Captured knowledge improves future project execution and builds organisational capability."
    ),
    (
        "Define benefit realisation.",
        ["Achieving the intended business benefits from the project",
         "Identify expected benefits before project starts",
         "Track benefit achievement during and after delivery",
         "Manage activities to maximise benefits",
         "Measure actual benefits delivered against expectations"],
        "Benefit realisation involves identifying expected benefits, tracking benefit achievement, managing activities to maximise benefits, and measuring actual benefits delivered. Many projects fail to realise full benefits; PM should focus on benefit delivery."
    ),
    (
        "What is return on investment (ROI)?",
        ["Measure of financial benefit relative to investment",
         "ROI = (Benefit - Cost) / Cost × 100%",
         "Projects should deliver positive ROI",
         "Used to justify and prioritise projects",
         "Higher ROI projects prioritised in portfolio selection"],
        "ROI measures profit generated relative to investment: (Benefit - Cost) / Cost × 100%. Projects should deliver positive ROI. ROI analysis helps prioritise projects and justify resource investment."
    ),
    (
        "Explain payback period.",
        ["Time until cumulative benefits equal the investment",
         "Shorter payback is generally preferable",
         "Simple and widely understood metric",
         "Does not account for time value of money",
         "Used alongside NPV and IRR for full financial assessment"],
        "Payback period is the duration until cumulative benefits equal investment. Shorter payback is generally preferable. Payback period is a simple metric but doesn't account for benefits after payback or time value of money."
    ),
    (
        "What is cost-benefit analysis?",
        ["Comparing all project costs against anticipated benefits",
         "Determines if project is financially worthwhile",
         "Helps justify project to stakeholders and sponsors",
         "Used to compare project alternatives",
         "Prioritises competing projects in portfolio selection"],
        "Cost-benefit analysis compares all project costs against anticipated benefits to determine if the project is worthwhile. Analysis helps justify the project, compare alternatives, and prioritise among competing projects."
    ),
    (
        "Define sensitivity analysis.",
        ["Assessing how changes in key variables affect outcomes",
         "Tests impact of schedule, cost, resource changes",
         "High-sensitivity variables require more careful management",
         "Informs contingency planning",
         "Helps identify most critical assumptions in the plan"],
        "Sensitivity analysis examines how changes in key assumptions (schedule, costs, resource availability) affect project outcome. High-sensitivity variables require more careful management and contingency planning."
    ),
    (
        "What is scenario planning?",
        ["Developing alternative plans for different future outcomes",
         "Best case, most likely, and worst case scenarios",
         "Develops associated strategies for each scenario",
         "Prepares PM for uncertainty",
         "Identifies key assumptions affecting outcomes"],
        "Scenario planning develops multiple potential outcomes (best case, most likely, worst case) and associated strategies. It helps prepare for uncertainty and identify key assumptions affecting project outcomes."
    ),
    (
        "Explain the role of PM in organisational change.",
        ["Supporting and managing transition to a new state",
         "Develops change management plan",
         "Trains staff on new processes or systems",
         "Communicates benefits and addresses resistance",
         "Ensures adoption and embedding of the change"],
        "When projects result in organisational change, the PM helps manage transition: developing a change management plan, training staff, communicating benefits, addressing resistance, and ensuring adoption of new ways of working."
    ),
    (
        "What is change management in project context?",
        ["Managing organisational and process transitions",
         "Awareness: why change is needed",
         "Training: how to work differently",
         "Support: removing adoption barriers",
         "Reinforcement: sustaining new ways of working"],
        "Change management encompasses supporting people through transitions to new ways of working. It includes awareness building, training, support (removing barriers), and reinforcement (sustaining new ways) to ensure successful adoption."
    ),
    (
        "Define critical success factors (CSF).",
        ["Essential conditions that must be present for project success",
         "Executive sponsorship and clear requirements",
         "Skilled team and effective communication",
         "Risk management and adequate resources",
         "Understanding CSFs focuses effort on what matters most"],
        "CSFs are critical factors that must be present for project success: executive sponsorship, clear requirements, skilled team, effective communication, risk management, adequate resources. Understanding CSFs helps focus effort on what matters most."
    ),
    (
        "What is a project constraint log?",
        ["Record of all limitations affecting project execution",
         "Documents budget, schedule, resource, and technical constraints",
         "Understanding constraints enables realistic planning",
         "Reviewed regularly as constraints change",
         "Shared with all project team members"],
        "A constraint log documents all factors limiting project execution: budget limits, schedule constraints, resource constraints, technical constraints, and organisational constraints. Understanding constraints helps develop realistic plans."
    ),
    (
        "Explain assumption log.",
        ["Record of assumptions underlying the project plan",
         "Documents key assumptions about team, technology, and vendors",
         "Invalid assumptions require plan reassessment",
         "Reviewed regularly throughout the project",
         "Prevents surprises from unvalidated assumptions"],
        "An assumption log documents key assumptions in the project plan. When assumptions prove invalid, project impact must be assessed and plans adjusted. Regular review prevents assumptions from becoming unmanaged risks."
    ),
    (
        "What is the concept of managing by walking around?",
        ["Direct observation of project activities",
         "PM regularly observes actual work and talks with team",
         "Identifies issues early before they escalate",
         "Maintains PM visibility and team morale",
         "Reveals problems formal status reports often miss"],
        "Managing by walking around (MBWA) involves the PM regularly observing actual work, talking with team members, identifying issues early, and maintaining visibility. MBWA often reveals problems that formal status reports miss."
    ),
    (
        "Define project retrospective.",
        ["Reflection on project results, process, and team performance",
         "Covers what was accomplished and what could improve",
         "Captures lessons for future projects",
         "Supports continuous improvement culture",
         "Similar to lessons learned but more frequent"],
        "A retrospective reflects on what was accomplished, what could improve, and what will be done differently. Regular retrospectives support continuous improvement and organisational learning."
    ),
    (
        "What is a project post-mortem?",
        ["Analysis conducted after project completion",
         "Examines what happened and why",
         "Captures learnings for future benefit",
         "Supports organisational knowledge growth",
         "Feeds process improvement initiatives"],
        "Post-mortem analysis after project completion examines what happened, why, what was learned, and how future projects can benefit. Post-mortems support organisational knowledge growth and process improvement."
    ),
    (
        "What is the strategic importance of Project Manager in organisation?",
        ["Drives project success and delivers business value",
         "Manages risk and uncertainty on behalf of the organisation",
         "Builds team capability and organisational knowledge",
         "Supports and enables organisational strategy execution",
         "Enables transformation and competitive advantage"],
        "The Project Manager is critical to organisational success by ensuring projects deliver intended value on time and within budget, managing risk and uncertainty, building team capability, supporting organisational strategy, and enabling organisational transformation and growth."
    ),
]


# ── 100 Defect Manager Q&A ────────────────────────────────────────────────
DEFECT_MANAGER_QA = [
    (
        "What is a defect in software testing?",
        ["Deviation from expected behaviour or specification",
         "Any flaw in code, design, or documentation",
         "Causes incorrect results or system failure",
         "Ranges from critical (blocking) to trivial (cosmetic)",
         "Identification is core to quality assurance"],
        "A defect (also called bug or issue) is any deviation from specified requirements or expected behaviour. It's a flaw in code, design, or documentation that causes incorrect results. Defects vary from critical (blocking functionality) to trivial (cosmetic issues). Defect identification is core to quality assurance."
    ),
    (
        "Define defect management.",
        ["Systematic process of capturing, tracking, and resolving defects",
         "Covers full defect lifecycle from report to closure",
         "Includes prioritisation, assignment, and verification",
         "Uses trend analysis for continuous improvement",
         "Provides quality visibility to stakeholders"],
        "Defect management is the systematic process of capturing, tracking, reporting, and resolving defects found during testing or production. It includes defect lifecycle management, prioritisation, assignment, resolution verification, and trend analysis."
    ),
    (
        "What are the main responsibilities of a Defect Manager?",
        ["Oversees the complete defect lifecycle",
         "Ensures defects are properly logged and tracked",
         "Facilitates defect resolution discussions",
         "Maintains defect metrics and generates reports",
         "Drives process improvement from defect data"],
        "A Defect Manager oversees the entire defect lifecycle: ensuring defects are properly logged, tracked, prioritised, assigned to resolution, verified after fixing, and closed. They maintain defect metrics, generate reports, and facilitate defect resolution discussions."
    ),
    (
        "Explain the defect lifecycle.",
        ["New → Assigned → Open/In Progress → Fixed → Verified → Closed",
         "Each stage has specific criteria and actions",
         "Defect can be Reopened if fix is insufficient",
         "Can be Deferred or Rejected at any stage",
         "Transition rules maintain data integrity"],
        "The defect lifecycle includes stages: New (reported) → Assigned → Open/In Progress (being investigated) → Fixed (resolution implemented) → Test/Verification (re-testing) → Closed (verified fixed) or Reopened (if still failing). Each stage has specific criteria and actions."
    ),
    (
        "What is defect severity?",
        ["Impact level of defect on system functionality",
         "Critical: system non-functional",
         "High: major feature impaired",
         "Medium: feature partially works",
         "Low/Trivial: minor or cosmetic impact"],
        "Defect severity measures how severely a defect impacts the system: Critical (system non-functional), High (major feature impaired), Medium (feature partially works), Low (minor impact), Trivial (cosmetic). Severity determines prioritisation and resource allocation."
    ),
    (
        "Define defect priority.",
        ["Urgency of fixing defect relative to business needs",
         "Critical: fix immediately",
         "High: fix in current sprint or release",
         "Medium: schedule normally",
         "Low: fix when resources allow"],
        "Defect priority indicates urgency of fixing relative to business value and deadlines. Priority considers business impact, not just technical severity. A cosmetic defect on a login page could have high priority due to brand impact."
    ),
    (
        "What is a defect report?",
        ["Detailed documentation of an identified defect",
         "Contains ID, title, description, and steps to reproduce",
         "Documents expected vs actual results",
         "Includes severity, priority, and environment details",
         "Attachments (screenshots, logs) speed resolution"],
        "A defect report documents: defect ID, title, description, steps to reproduce, expected vs actual results, severity, priority, environment details (OS, browser, version), attachments (screenshots, logs), and date reported. Complete reports enable faster resolution."
    ),
    (
        "Explain reproducibility in defects.",
        ["Ability to consistently recreate the defect",
         "Reproducible defects are easier to diagnose and fix",
         "Intermittent defects are harder and need more investigation",
         "Steps to reproduce must be precise and complete",
         "Environment-specific defects may need environment replication"],
        "Reproducibility indicates whether a defect can be reliably recreated by following the steps to reproduce. Reproducible defects are easier to fix; intermittent defects are harder to diagnose and resolve, often requiring environment-specific investigation."
    ),
    (
        "What is defect root cause?",
        ["Underlying fundamental reason the defect was introduced",
         "Different from the immediate (surface) cause",
         "Found through Root Cause Analysis (RCA)",
         "Enables permanent fix and future prevention",
         "Examples: missing validation, incorrect logic, misunderstood requirement"],
        "Root cause is the fundamental reason a defect exists, often different from the immediate cause. Root cause analysis (RCA) identifies why the defect occurred, enabling permanent fixes and prevention of similar defects."
    ),
    (
        "Define defect triage.",
        ["Process of reviewing, prioritising, and assigning new defects",
         "Validates whether report is an actual defect",
         "Assesses severity and priority consistently",
         "Assigns defect to appropriate team or individual",
         "Daily triage meetings manage defect backlog"],
        "Defect triage reviews newly reported defects, validates they're actual defects (not user error), assesses severity/priority, assigns to appropriate team/individual, and schedules for resolution. Triage meetings typically occur daily to manage defect backlog."
    ),
    (
        "What is a defect register/database?",
        ["Centralised repository of all project defects",
         "Tools: JIRA, Azure DevOps, HP Quality Center",
         "Stores status, tracking history, and associated info",
         "Enables trend analysis, reporting, and compliance",
         "Single source of truth for defect status"],
        "A defect database (tools like JIRA, Azure DevOps, HP QC) stores all defect records with status, tracking history, and associated information. The database enables visibility, trend analysis, reporting, and process compliance."
    ),
    (
        "Explain defect duplicate detection.",
        ["Identifying multiple reports of the same issue",
         "Prevents redundant work on the same problem",
         "Consolidate duplicates to one master record",
         "Keep links between duplicate records",
         "Duplicates distort metrics if not consolidated"],
        "Duplicate detection prevents redundant work on the same problem reported multiple times. Defect Managers should consolidate duplicates, keeping one master record with links to duplicates. Duplicates distort defect metrics if not consolidated."
    ),
    (
        "What is defect status?",
        ["Current state of defect in its lifecycle",
         "States: New, Assigned, Open, Fixed, Test, Closed",
         "Also: Reopened, Deferred, or Rejected",
         "Status visibility helps team understand work remaining",
         "Enables prioritisation and workload management"],
        "Defect status indicates where the defect is in its lifecycle: New, Assigned, Open, Fixed, Test, Closed, Reopened, Deferred, or Rejected. Status visibility helps the team understand work remaining and enables prioritisation."
    ),
    (
        "Define defect assignment.",
        ["Designation of responsible party for defect resolution",
         "Based on expertise, availability, and skill match",
         "Must be clearly communicated to assignee",
         "Prevents work from falling through gaps",
         "Tracked until assignee accepts responsibility"],
        "Defect assignment specifies who is responsible for fixing the defect. Assignment should be based on expertise, availability, and skill match. Clear assignment prevents work from falling through the cracks."
    ),
    (
        "What is defect verification?",
        ["Re-testing to confirm the defect is fixed",
         "Also called re-testing or confirmation testing",
         "Executes exact steps that originally exposed the defect",
         "Defect reopened if issue persists",
         "Closed only when fix is confirmed successful"],
        "Defect verification involves executing test steps after a fix is implemented to confirm the defect is resolved. If the defect still exists, it is reopened; if confirmed fixed, it is closed. Verification is critical before declaring a defect resolved."
    ),
    (
        "Explain regression related to defects.",
        ["Ensures a fix doesn't introduce new defects",
         "Fixed code may break existing functionality",
         "Regression suite covers affected and dependent areas",
         "Systematic regression testing prevents surprises",
         "Critical after every code change"],
        "Regression testing after a defect fix ensures the fix didn't introduce new defects or break existing functionality. Regression is a common risk; systematic regression testing prevents it."
    ),
    (
        "What is defect density?",
        ["Number of defects per unit of code or functionality",
         "Measured as defects per KLOC (1000 lines of code)",
         "Also per function point or test case",
         "Comparison across versions reveals quality trends",
         "High density indicates quality issues in a module"],
        "Defect density measures quality: defects per 1000 lines of code (KLOC), defects per function point, or defects per test case. Density comparison across projects or versions identifies quality trends."
    ),
    (
        "Define defect escape.",
        ["Defect not caught before production release",
         "Represents a failure of the testing process",
         "Escapes damage customer experience and trust",
         "Analysed to improve test coverage and techniques",
         "Key metric for testing effectiveness"],
        "A defect escape is a bug that wasn't found during testing and made it to production. Escapes represent testing failures. Analysing escapes helps improve test coverage, techniques, and defect detection effectiveness."
    ),
    (
        "What is defect detection rate?",
        ["Percentage of total defects found during testing",
         "(Defects Found / Total Defects) × 100%",
         "Higher rate indicates more effective testing",
         "Compared across projects to identify trends",
         "Low detection rate means more escapes likely"],
        "Detection rate measures testing effectiveness: (Defects Found / Total Defects) × 100%. Higher detection rates indicate better testing. Comparing detection rates across projects identifies testing quality trends."
    ),
    (
        "Explain defect distribution.",
        ["Breakdown of defects by category or component",
         "By module: shows which areas have most defects",
         "By type: logic errors vs syntax vs design",
         "By phase: when defects are introduced and found",
         "Distribution reveals patterns and improvement opportunities"],
        "Defect distribution analysis shows where defects concentrate: by module, by type (logic, syntax, design), or by phase (requirements vs testing). Distribution reveals patterns and improvement opportunities."
    ),
    (
        "What is a defect trend analysis?",
        ["Examining defect patterns over time",
         "Are defects increasing or decreasing?",
         "Which components consistently have defects?",
         "Which types (logic, design) are most common?",
         "Trends reveal systemic quality issues"],
        "Trend analysis reviews historical defect data to identify patterns: Are defects increasing or decreasing? Which components consistently have defects? Are certain types more common? Trends reveal systemic quality issues."
    ),
    (
        "Define defect correlation.",
        ["Relationship between defects or related issues",
         "Multiple defects may share a root cause",
         "Defects from specific coding pattern",
         "Defects in related or dependent modules",
         "Correlations enable more efficient root cause analysis"],
        "Defects may correlate: multiple defects from same root cause, defects from specific coding patterns, or defects in related modules. Identifying correlations enables more efficient root cause analysis and prevention."
    ),
    (
        "What is defect criticality assessment?",
        ["Determining business impact of the defect",
         "Does it block critical business processes?",
         "Does it cause data loss or corruption?",
         "Is sensitive or regulated data at risk?",
         "Business impact, not just technical severity, drives priority"],
        "Criticality assessment evaluates how the defect affects business: Does it block critical processes? Does it cause data loss? Is sensitive data at risk? Business impact, not just technical severity, drives prioritisation."
    ),
    (
        "Explain defect resolution strategy.",
        ["Approach to fixing the identified defect",
         "Fix immediately: critical blocking issues",
         "Schedule fix: planned in next sprint or release",
         "Defer: lower priority pushed to later release",
         "Workaround or Won't Fix for low-value defects"],
        "Resolution strategies include: Fix immediately (critical issues), Schedule fix (planned work), Defer to later release (lower priority), Workaround (temporary solution), or Won't fix (cost-benefit unfavourable). Strategy depends on severity, priority, and resources."
    ),
    (
        "What is a defect log?",
        ["Simple record of defects and their current status",
         "Less formal than a full defect database",
         "Useful for small projects",
         "Lists defect description, status, severity, owner",
         "Can serve as backup to formal defect system"],
        "A defect log is basic documentation listing defects, status, severity, and key details. Less formal than a full defect database, logs are useful in small projects or as backup to a formal system."
    ),
    (
        "Define defect metrics.",
        ["Quantitative measures related to defect management",
         "Count, severity distribution, priority distribution",
         "Resolution time, detection rate, escape rate",
         "Density and trend analysis over time",
         "Metrics enable visibility, comparison, and improvement"],
        "Defect metrics include: count, severity distribution, priority distribution, resolution time, detection rate, escape rate, density, and trend analysis. Metrics enable visibility, comparison, and continuous improvement."
    ),
    (
        "What is a defect SLA (Service Level Agreement)?",
        ["Time commitment for defect resolution by severity",
         "Critical: resolve within 4 hours",
         "High: resolve within 1 business day",
         "Medium: resolve within 3 business days",
         "Low: resolve within 1 week"],
        "Defect SLAs define expected resolution times by severity: Critical (4 hours), High (1 day), Medium (3 days), Low (1 week). SLAs set stakeholder expectations and hold the team accountable for timely resolution."
    ),
    (
        "Explain defect communication.",
        ["Notifying stakeholders about defect status",
         "Developers informed when defects are assigned to them",
         "Testers notified when fixes are ready for verification",
         "PMs receive regular defect status reports",
         "Critical issues escalated immediately"],
        "Defect communication includes: informing developers when defects are assigned, notifying testers when fixes are ready for verification, reporting defect status to project managers, and escalating critical issues. Clear communication expedites resolution."
    ),
    (
        "What is a known issue?",
        ["Confirmed defect with no immediate resolution planned",
         "Decision made by business to defer or accept the defect",
         "Documented for stakeholder awareness",
         "Users provided workaround if available",
         "Reviewed for resolution in future releases"],
        "A known issue is a confirmed defect for which no immediate fix is planned (due to cost, complexity, or business decision). Known issues are documented so stakeholders understand limitations. Users may require workarounds."
    ),
    (
        "Define defect waiver.",
        ["Formal documentation accepting release despite known defect",
         "Business decision balancing risk vs benefit",
         "Typically for lower-severity defects",
         "Requires sign-off from stakeholders",
         "Workaround should exist for waivers"],
        "A defect waiver is formal documentation acknowledging a defect exists but approving release despite it. Waivers are business decisions, typically for lower-severity issues where benefits outweigh risks or workarounds exist."
    ),
    (
        "What is the role of Defect Manager in quality assurance?",
        ["Ensures systematic defect capture and resolution",
         "Maintains quality standards throughout project",
         "Generates insights from defect data",
         "Facilitates root cause analysis",
         "Drives process improvements to prevent recurrence"],
        "The Defect Manager ensures defects are systematically captured, tracked, resolved, and prevented. They maintain quality standards, generate insights from defect data, facilitate root cause analysis, and drive process improvements."
    ),
    (
        "Explain root cause analysis (RCA).",
        ["Systematic investigation of why a defect occurred",
         "Asks 'why' iteratively to find fundamental cause",
         "Distinguishes symptoms from root causes",
         "Enables permanent fix not just symptom relief",
         "Prevents recurrence of similar defects"],
        "RCA involves asking 'why' iteratively to find the fundamental cause: Why did this defect occur? (developer didn't understand requirement) Why? (requirement wasn't clearly communicated) Why? (no requirement review process). RCA prevents recurrence."
    ),
    (
        "What is the 5-Why technique?",
        ["RCA method asking 'why' five times",
         "Drills from symptom to root cause",
         "Simple but effective for systemic issues",
         "Each 'why' answers the previous 'why'",
         "Prevents addressing symptoms instead of causes"],
        "The 5-Why technique involves asking 'why' repeatedly (usually five times) to drill down from symptom to root cause. Simple but effective for finding underlying issues rather than surface symptoms."
    ),
    (
        "Define Pareto principle in defect context.",
        ["80% of defects come from 20% of causes",
         "Identifying the vital 20% maximises improvement ROI",
         "Focus effort on highest-impact causes",
         "Used to prioritise defect prevention activities",
         "Data-driven application using defect trend analysis"],
        "The Pareto principle suggests 80% of defects come from 20% of causes. Identifying this vital 20% and addressing it yields maximum quality improvement return on effort invested."
    ),
    (
        "What is defect classification?",
        ["Categorisation of defects by type or source",
         "By type: logic, syntax, design, requirement, configuration",
         "By source: coding error, misunderstood requirement, test issue",
         "By component: database, UI, API, integration",
         "Classification helps identify patterns and areas for improvement"],
        "Defects are classified by type (logic, syntax, design, requirement, configuration), source (requirement misunderstanding, coding error), or component (database, UI, API). Classification helps identify patterns and improvement areas."
    ),
    (
        "Explain defect slippage.",
        ["Defect fix not completed for planned release",
         "Resolution pushed to a later version",
         "Must be formally documented and communicated",
         "Multiple slippages indicate planning or execution problems",
         "Each slippage should be analysed for cause"],
        "Slippage occurs when planned fixes aren't completed for a release, pushing resolution to a later version. Slippage must be formally documented, communicated, and tracked. Multiple slippages indicate unrealistic planning or execution problems."
    ),
    (
        "What is a defect bounce-back?",
        ["Defect closed by developer as 'not a bug'",
         "Developer disputes that it is actually a defect",
         "Defect Manager or QA lead reviews disputed cases",
         "Product Owner may be required to adjudicate",
         "High bounce-back rate indicates poor defect reporting"],
        "Bounce-back occurs when a developer disputes the defect claim or closes it as 'not a bug' (user error, expected behaviour). Disputed defects should be reviewed by the Defect Manager, QA lead, or Product Owner for resolution."
    ),
    (
        "Define defect reopening.",
        ["Returning a defect to open status after being closed",
         "Happens when verification testing shows issue still exists",
         "Or when fix introduced a new problem",
         "High reopening rate indicates poor fix quality",
         "Reopened defects require developer reassignment"],
        "A defect is reopened if verification testing after the fix finds the issue still exists or the fix created problems. Reopened defects indicate incomplete or incorrect fixes. High reopening rates suggest process or quality issues."
    ),
    (
        "What is a defect batch?",
        ["Group of related defects addressed together",
         "Groups by component, developer, or release",
         "Can speed resolution through focused effort",
         "Risk: may delay critical individual fixes",
         "Balance batching with priority needs"],
        "Defects may be batched for efficiency: all database defects in one fix, or all UI defects in another. Batching can speed resolution but may delay critical individual fixes. Balance is needed."
    ),
    (
        "Explain defect criticality levels.",
        ["Severity classification for urgency of handling",
         "P0/P1: emergency — requires immediate all-hands response",
         "P2/P3: urgent — fix in current sprint",
         "P4: routine — schedule normally",
         "Criticality triggers specific escalation paths"],
        "Criticality levels (P0/P1 for emergency, P2/P3 for urgent, P4 for routine) indicate urgency. Highest criticality requires immediate escalation, all-hands-on-deck response, and rapid resolution."
    ),
    (
        "What is a defect impact analysis?",
        ["Assessment of defect effects across the system",
         "Evaluates not just immediate but downstream effects",
         "Does it corrupt data used elsewhere?",
         "Does it break dependent functionality?",
         "Full impact understanding informs resolution strategy"],
        "Impact analysis evaluates not just the immediate problem but how the defect affects other areas: Does it corrupt data used elsewhere? Does it break dependent functionality? Understanding full impact informs the resolution strategy."
    ),
    (
        "Define defect lifecycle stage transitions.",
        ["Rules governing valid defect status changes",
         "Defect can only be closed if fixed and verified",
         "Can only reopen if previously closed",
         "Invalid transitions prevented by tool workflow",
         "Clear rules maintain data integrity"],
        "Stage transitions have specific rules: a defect can only be closed if fixed and verified, can only reopen if closed, etc. Clear transition rules prevent invalid state changes and maintain data integrity."
    ),
    (
        "What is a defect prioritisation matrix?",
        ["Tool for determining defect priority objectively",
         "Plots severity vs probability/frequency",
         "High severity + high frequency = top priority",
         "Low severity + low frequency = lowest priority",
         "Provides consistent, data-driven prioritisation"],
        "A prioritisation matrix plots severity (y-axis) against probability/frequency (x-axis), creating quadrants for priority assignment. High-severity + high-frequency defects get top priority; low-severity + low-frequency get the lowest priority."
    ),
    (
        "Explain defect resolution tracking.",
        ["Monitoring progress of individual defect fixes",
         "Tracks when developer starts work",
         "Monitors milestones and expected completion",
         "Identifies blockers early",
         "Enables escalation when SLAs are at risk"],
        "Resolution tracking monitors: when developer starts work, milestones of progress, expected completion, and any blockers. Regular tracking enables early identification of delays and expedites resolution."
    ),
    (
        "What is a defect postmortem?",
        ["Analysis of a defect after it has been resolved",
         "How did it occur? How was it detected?",
         "Was the fix high quality?",
         "Can recurrence be prevented?",
         "Lessons documented for team learning"],
        "Postmortem review after defect closure analyses: how defect occurred, how it was detected, quality of fix, whether prevention is possible, and lessons learned. Postmortems on critical defects prevent recurrence."
    ),
    (
        "Define defect leak/escape analysis.",
        ["Investigation of defects that reached production",
         "Why did testing miss them?",
         "Was test coverage insufficient?",
         "Were test conditions inadequate or wrong?",
         "Findings improve testing process to prevent future escapes"],
        "When production defects are found, analysis determines why testing missed them: insufficient test coverage, inadequate test conditions, or poor test design. Leak analysis improves testing processes."
    ),
    (
        "What is severity vs. priority trade-off?",
        ["Balance between technical impact and business urgency",
         "Severity = technical impact on system",
         "Priority = business urgency to fix",
         "High-severity defect may have lower priority if rarely triggered",
         "Medium-severity defect affecting many users gets higher priority"],
        "Severity reflects technical impact; priority reflects business urgency. A high-severity, low-frequency defect affecting few users gets lower priority than a medium-severity defect affecting many users. The Defect Manager must balance both dimensions."
    ),
    (
        "Explain defect escalation.",
        ["Raising critical defect to higher authority",
         "Triggered when normal resolution is too slow",
         "Triggers: critical blocking progress, SLA breach, deadlock",
         "Escalation path must be defined in advance",
         "Defect Manager owns the escalation process"],
        "Escalation moves critical defects to higher management when normal resolution processes are too slow. Escalation triggers: critical defect blocking progress, resolution time exceeding SLA, or deadlock in resolution approach."
    ),
    (
        "What is a defect workaround?",
        ["Temporary procedure to avoid defect impact",
         "Not a permanent solution",
         "Enables users or business to continue working",
         "Documented and communicated to affected users",
         "Removed once permanent fix is deployed"],
        "A workaround is a temporary procedure users can follow to avoid defect impact while the permanent fix is developed. Workarounds prevent complete blocking but are not permanent solutions. Users should be documented on workarounds."
    ),
    (
        "Define defect churn.",
        ["Excessive opening and closing cycles without stable resolution",
         "Defects repeatedly opened, closed, reopened by different people",
         "Indicates process or communication breakdown",
         "Increases resolution time and team frustration",
         "Defect Manager must identify and resolve churn patterns"],
        "Churn occurs when defects are repeatedly opened, closed, and reopened without a stable fix. Churn indicates process or communication breakdown and increases resolution time and frustration."
    ),
    (
        "What is the role of Defect Manager in testing phases?",
        ["Different phases require different defect management approaches",
         "Unit testing: ensure developers verify all fixes",
         "Integration testing: track cross-component issues",
         "System testing: manage high-volume defect flow",
         "UAT: prioritise business-impacting defects"],
        "In unit testing, the Defect Manager ensures developers verify fixes. In integration testing, they track cross-component issues. In system testing, they manage high-volume defect flow. In UAT, they prioritise business-impacting defects."
    ),
    (
        "Explain defect aging.",
        ["Tracking how long defects remain unresolved",
         "Aging trends reveal resolution bottlenecks",
         "Long-aging defects should be escalated",
         "Target: close most defects within SLA",
         "Aging report is key Defect Manager tool"],
        "Defect aging analysis shows how long defects stay open before closure. Aging trends reveal if resolution is getting slower or bottlenecks exist. Long-aging defects should be escalated. The target is to close most defects within their SLA."
    ),
    (
        "What is a defect dashboard?",
        ["Visual display of defect status and key metrics",
         "Shows total open defects by severity and priority",
         "Displays distribution by component and aging trends",
         "Provides quick visibility on defect health",
         "Helps identify bottlenecks requiring action"],
        "Defect dashboards display key metrics: total open defects, defect distribution by severity/priority/component, aging trends, and resolution trends. Dashboards provide quick visibility on defect health and bottlenecks."
    ),
    (
        "Define defect categorisation by source.",
        ["Classifying defects by where the issue originated",
         "Requirements: misunderstood or missing requirement",
         "Design: flawed architecture or design decision",
         "Development: coding error",
         "Testing or Environmental: issue in test or environment itself"],
        "Source categorisation includes: Requirements (misunderstood requirement), Design (flawed architecture), Development (coding error), Testing (test issue), or Environmental (environment-specific problem). Source analysis reveals quality issues in specific areas."
    ),
    (
        "What is defect correlation analysis?",
        ["Finding relationships between multiple defects",
         "Multiple defects from the same root cause",
         "Defects from same developer, module, or component",
         "Defects occurring after specific types of change",
         "Correlations enable efficient fixing and prevention"],
        "Correlation analysis identifies: multiple defects from same root cause, defects from same developer/module, or defects following patterns (all on specific OS, after specific change). Correlations enable efficient fixing and prevention."
    ),
    (
        "Explain defect prevention strategies.",
        ["Requirements reviews catch issues before development",
         "Code reviews: peer checking reduces coding errors",
         "Static analysis: automated code checking",
         "TDD: testing before coding prevents defects",
         "Training: skilled developers introduce fewer defects"],
        "Prevention includes: requirements reviews (catch issues early), code reviews (peer review catches errors), static analysis (automated code checking), TDD (testing before coding), pair programming, and training (skilled developers make fewer mistakes)."
    ),
    (
        "What is a defect control plan?",
        ["Document defining the defect management processes",
         "Specifies how defects are reported and triaged",
         "Defines severity/priority scale and SLAs",
         "Documents assignment and resolution workflow",
         "Ensures consistent, transparent defect management"],
        "The control plan specifies: how defects are reported, severity/priority definitions, assignment process, resolution timelines, verification process, and closure criteria. It ensures consistent, transparent defect management."
    ),
    (
        "Define defect investigation.",
        ["Detailed analysis to understand a specific defect",
         "Involves reproducing, analysing logs and data",
         "Examines code and tests theories",
         "Ultimately determines root cause",
         "Enables accurate fix and full impact understanding"],
        "Investigation involves reproducing the defect, analysing logs and data, examining code, testing theories, and ultimately determining root cause. Thorough investigation enables accurate fixes and understanding of impact."
    ),
    (
        "What is a defect summary report?",
        ["High-level overview of defect status and trends",
         "Total defects by severity and trend over time",
         "Aging analysis and resolution rate",
         "Designed for executives and managers",
         "Supports release readiness decision-making"],
        "Summary reports provide executives and managers a quick view: total defects by severity, trends over time, aging analysis, and resolution rate. Summary reports facilitate decision-making about quality status and release readiness."
    ),
    (
        "Explain defect management in DevOps.",
        ["Rapid defect handling in continuous deployment environments",
         "Defect tools integrate with CI/CD pipelines",
         "Automated testing feeds defects automatically",
         "Resolution must be fast — no time for lengthy triage",
         "Defect Manager must adapt processes for DevOps speed"],
        "In DevOps, defects must be managed rapidly since deployments are frequent. Defect tools integrate with CI/CD pipelines, automated testing feeds defects, and resolution must be fast. The Defect Manager must adapt processes for DevOps speed."
    ),
    (
        "What is a defect handoff?",
        ["Transferring defect ownership between team members",
         "From tester (reporter) → developer (fixer) → QA (verifier)",
         "Clear handoff communication prevents misunderstandings",
         "Defect Manager ensures smooth transitions",
         "Prevents defects stalling between team members"],
        "Handoff occurs when a defect moves from tester (reporter) to developer (fixer) to QA (verifier). Clear handoff communication prevents misunderstandings. The Defect Manager ensures smooth transitions and prevents defects from stalling at handoffs."
    ),
    (
        "Define defect documentation standards.",
        ["Guidelines for consistent, complete defect reports",
         "Required fields: title, steps to reproduce, expected/actual results",
         "Consistent format reduces ambiguity",
         "Precision requirements for environment details",
         "Good standards reduce resolution time"],
        "Standards specify: required fields (title, description, steps to reproduce, expected/actual results), information to include (environment, attachments), format consistency, and precision requirements. Good standards reduce ambiguity and enable faster resolution."
    ),
    (
        "What is a defect metrics dashboard?",
        ["Real-time display of defect KPIs",
         "Shows defect arrival rate (new defects per day)",
         "Displays resolution rate (defects fixed per day)",
         "Open count trend and severity distribution",
         "Enables data-driven defect management"],
        "Metrics dashboards show: defect arrival rate (new defects per day), resolution rate (defects fixed per day), open count trend, severity distribution, aging distribution, and escape rate. Dashboards enable data-driven management."
    ),
    (
        "Explain defect estimation.",
        ["Predicting final defect count and resolution effort",
         "Uses historical data and detection rate patterns",
         "Estimates likely final defect count",
         "Projects resources and time needed",
         "Supports release planning and scheduling"],
        "Using historical data, managers estimate: likely final defect count (based on detection rate), resolution time needed, and resources required. Estimation helps plan testing phases and release dates."
    ),
    (
        "What is a critical defect task force?",
        ["Dedicated team assembled for urgent defect resolution",
         "Developers, QA, and management work together",
         "Works intensively until defect is resolved",
         "Bypasses normal workflow for speed",
         "Used for production-impacting P0/P1 defects"],
        "For critical production defects, a dedicated task force may be assembled with developers, QA, and management to achieve rapid resolution. The task force works intensively until the defect is fixed and verified."
    ),
    (
        "Define defect prevention metrics.",
        ["Measures of how effectively defects are prevented",
         "Requirements review effectiveness: defects caught in review",
         "Code review effectiveness: errors caught before testing",
         "Static analysis catch rate",
         "Percentage of defects prevented before formal testing"],
        "Prevention metrics include: requirements review effectiveness (% defects caught in review vs. testing), code review effectiveness, static analysis effectiveness, and percentage of defects prevented before testing."
    ),
    (
        "What is a defect policy?",
        ["Formal rules governing how defects are handled",
         "Defines what qualifies as a defect",
         "Sets severity/priority scale and SLAs",
         "Specifies documentation requirements",
         "Defines escalation procedures and change control for fixes"],
        "Defect policy defines: what qualifies as a defect, severity/priority scale, SLAs for resolution, required documentation, change control for fixes, and escalation procedures. Policy ensures consistency across projects and teams."
    ),
    (
        "Explain defect communication to stakeholders.",
        ["Keeping all interested parties informed",
         "Daily defect status reports during active testing",
         "Weekly trend reports for project managers",
         "Monthly metrics presentations for leadership",
         "Immediate escalation of critical or blocking issues"],
        "Stakeholder communication includes: daily defect status reports, weekly trend reports, monthly metrics presentations, and immediate escalation of critical issues. Communication frequency depends on phase and criticality."
    ),
    (
        "What is a defect screenshot/mock-up?",
        ["Visual documentation of the defect",
         "Shows exact error message or unexpected behaviour",
         "Helps developers understand issue faster than text alone",
         "Prevents misinterpretation of the defect",
         "Reduces back-and-forth and resolution time"],
        "Screenshots or recordings showing the defect help developers understand the exact issue faster than text description. Visual documentation also prevents misinterpretation and reduces resolution time."
    ),
    (
        "Define defect environment tracking.",
        ["Recording conditions under which defect occurs",
         "Includes OS, browser version, RAM, network",
         "Database version and third-party software versions",
         "Helps reproduce defects reliably",
         "Identifies if issue is environment-specific"],
        "Defect environment information includes: OS, browser version, RAM, network, database version, and third-party software. Environment details help reproduce defects and identify if the issue is environment-specific."
    ),
    (
        "What is a defect regression suite?",
        ["Tests to verify no new defects introduced by a fix",
         "Covers directly affected code",
         "Covers dependent and related functionality",
         "Areas likely affected by the code change",
         "Critical after every fix deployment"],
        "After a fix, regression tests ensure the fix doesn't break existing functionality. The regression suite should cover: directly affected code, dependent functionality, and areas likely affected by code changes."
    ),
    (
        "Explain defect batching strategy.",
        ["Grouping defects for efficient processing",
         "By component: all database fixes together",
         "By priority: critical fixes first",
         "By developer: each developer fixes their code",
         "Batching improves efficiency if defects stay on schedule"],
        "Batching strategy might group: by component (all database fixes together), by priority (all critical fixes first), by developer (each developer fixes their code), or by release. Batching improves efficiency if defects stay on schedule."
    ),
    (
        "What is a defect tracking tool vs. spreadsheet?",
        ["Dedicated tools are preferable for professional environments",
         "Spreadsheets work only for very small projects",
         "Tools provide workflow automation and real-time reporting",
         "Integration with development tools (JIRA, Azure DevOps)",
         "Tools scale; spreadsheets do not"],
        "Spreadsheets work for small projects but lack automation, reporting, and scalability. Dedicated tools (JIRA, Azure DevOps, QC) provide: workflow automation, real-time reporting, integration with development, and scalability. Tools are preferable for professional environments."
    ),
    (
        "Define defect severity assessment.",
        ["Process of consistently determining impact level",
         "Considers functional impact: does it block users?",
         "Data impact: is data lost or corrupted?",
         "Scope: how many users are affected?",
         "Workaround availability reduces effective severity"],
        "Assessment considers: functional impact (does it block users?), data impact (is data lost/corrupted?), scope (how many users affected?), and workaround availability. Assessment should be consistent using defined criteria."
    ),
    (
        "What is a defect triage checklist?",
        ["Standard list of items to verify when triaging defects",
         "Is it actually a defect or user error?",
         "What is the severity? Priority?",
         "Who should fix it? Is there a duplicate?",
         "Does it need more information before assignment?"],
        "Triage checklist includes: Is it actually a defect? What's the severity? What's the priority? Who should fix it? Is there a duplicate? Does it need more information? Using checklists ensures consistent, thorough triage."
    ),
    (
        "Explain defect impact on release decision.",
        ["Defect data drives go/no-go decisions",
         "Number of open critical/high defects is key factor",
         "Escape risk assessment for remaining open defects",
         "Customer impact if released with open defects",
         "Remediation plans for known issues post-release"],
        "Release decisions consider: number of open critical/high defects, escape risk assessment, customer impact if released, and remediation plans. Defect data enables informed release decisions and risk management."
    ),
    (
        "What is a defect contingency plan?",
        ["Plan for handling unresolved defects at release time",
         "Identifies which defects are acceptable to release with",
         "Identifies which require last-minute critical path fixes",
         "Identifies which justify a release delay",
         "Prepares team for various release scenarios"],
        "Contingency planning identifies: which defects are acceptable to release with (known issues), which require last-minute fixes (critical path items), and which justify release delays (blocking issues). Plans prepare the team for release scenarios."
    ),
    (
        "Define defect metrics baseline.",
        ["Reference point for comparing future defect performance",
         "Establishes initial detection rate, escape rate, resolution time",
         "Future metrics compared to baseline show improvement or degradation",
         "Baselines vary by project type and size",
         "Essential for measuring process improvement impact"],
        "Baseline establishes initial defect metrics (detection rate, escape rate, resolution time) for a project/product. Future metrics compared to the baseline show improvement or degradation trends. Baselines vary by project characteristics."
    ),
    (
        "What is defect velocity?",
        ["Rate at which defects are resolved over time",
         "Measured as defects resolved per sprint or week",
         "Stable velocity = predictable resolution capacity",
         "Declining velocity may indicate complexity or resource issues",
         "Used for release planning and capacity management"],
        "Defect velocity measures how many defects per sprint/week are resolved. Stable velocity indicates predictable resolution capacity. Declining velocity might indicate increasing complexity, resource issues, or process problems."
    ),
    (
        "Explain defect open rate analysis.",
        ["Examining why defects remain open",
         "Which severity or components stay open longest?",
         "What is blocking resolution?",
         "Missing information, blocked developer, or dependency?",
         "Analysis identifies systemic resolution obstacles"],
        "Open rate analysis identifies: what severity/component defects tend to remain open longest, what's blocking resolution (missing information, blocked developer, dependency issues), and what changes could expedite closure."
    ),
    (
        "What is a defect assignment strategy?",
        ["Method for allocating defects to the right developers",
         "Auto-assign by component owner",
         "Manager assigns based on skill and availability",
         "Developer picks from prioritised backlog",
         "Strategy should balance workload and expertise matching"],
        "Assignment strategies include: auto-assign by component owner, manager assigns based on skill/availability, developer picks from backlog, or priority-based auto-assignment. Strategy should balance workload and expertise matching."
    ),
    (
        "Define defect resolution effectiveness.",
        ["Measure of how well defect fixes work first time",
         "Percentage of fixes that work without rework",
         "Average time from assignment to resolution",
         "Reopening rate as key quality indicator",
         "High effectiveness indicates quality development process"],
        "Effectiveness tracks: percentage of fixes that work first time (vs. needing rework), average time to resolution, and reopening rate. High effectiveness indicates a quality development process; low effectiveness suggests process issues."
    ),
    (
        "What is a defect reduction plan?",
        ["Strategy for decreasing defect volume over time",
         "Based on defect trend analysis findings",
         "May include intensive testing or code review emphasis",
         "Developer training and automated testing expansion",
         "Requirements review enhancement for earlier detection"],
        "A reduction plan might include: intensive testing, code review emphasis, developer training, process improvements, automated testing expansion, or requirements review enhancement. Plans should be based on defect trend analysis."
    ),
    (
        "Explain defect categorisation by phase introduced.",
        ["Tracking when in the project the defect was introduced",
         "Requirements phase: misunderstood specification",
         "Design phase: flawed architecture decision",
         "Development phase: coding error",
         "Reveals which phases introduce most defects for targeted improvement"],
        "Categorising defects by introduction phase (requirements, design, development, testing) reveals which phases introduce most defects, enabling targeted process improvement in highest-impact areas."
    ),
    (
        "What is a defect cost-benefit analysis?",
        ["Evaluating whether fixing a defect is worthwhile",
         "Cost of fix: developer time, risk, testing effort",
         "Benefit: reduced support, improved UX, prevented failures",
         "Some low-priority defects may not justify fix cost",
         "Business decision made with stakeholder input"],
        "Analysis compares: cost of fixing (developer time, risk, testing) vs. benefit (reduced support calls, improved user experience, prevented failures). Some low-priority defects may not justify fix cost."
    ),
    (
        "Define defect management process improvement.",
        ["Enhancing defect management effectiveness continuously",
         "Faster triage: reducing time from report to assignment",
         "Better classification enabling faster diagnosis",
         "Integrated tools improving visibility across teams",
         "Clearer SLAs improving predictability and accountability"],
        "Improvements might include: faster triage (reducing time to assignment), better classification (enabling faster diagnosis), integrated tools (improving visibility), or clearer SLAs (improving predictability). Continuous improvement increases efficiency."
    ),
    (
        "What is defect management training?",
        ["Educating team members on defect processes",
         "Covers what constitutes a valid defect",
         "How to write effective, complete defect reports",
         "Severity and priority assessment guidance",
         "Tool usage and defect management standards"],
        "Training covers: what constitutes a defect, how to write effective defect reports, severity/priority assessment, defect lifecycle, tool usage, and defect management standards. Good training improves defect data quality."
    ),
    (
        "Explain defect automation.",
        ["Using tools to automate defect-related processes",
         "Automated testing detects defects automatically",
         "Automated workflow moves defects through lifecycle",
         "Automated notifications alert relevant parties",
         "Automated reporting generates dashboards and metrics"],
        "Automation includes: automated testing (detects defects), automated defect workflow (moves through lifecycle automatically), automated notifications (alerts relevant parties), and automated reporting (generates reports). Automation increases efficiency and consistency."
    ),
    (
        "What is a defect audit?",
        ["Review of defect management process compliance",
         "Verifies defects are properly classified",
         "Checks SLAs are being met consistently",
         "Ensures documentation is complete and accurate",
         "Identifies process compliance issues and improvement opportunities"],
        "Audits verify: defects are properly classified, SLAs are being met, documentation is complete, and process is being followed consistently. Audits identify process compliance issues and improvement opportunities."
    ),
    (
        "Define defect prioritisation in crisis.",
        ["Rapid prioritisation under time pressure",
         "Focus on what is blocking critical processes",
         "What is affecting the most users right now?",
         "What has the simplest and safest fix?",
         "Crisis prioritisation requires fast, decisive decision-making"],
        "In crisis (production outage), rapid prioritisation focuses on: what's blocking critical processes, what's affecting most users, and what has the simplest fix. Crisis prioritisation differs from routine prioritisation and requires quick decision-making."
    ),
    (
        "What is a defect root cause knowledge base?",
        ["Repository of common defect causes and proven solutions",
         "Documents common root causes in your systems",
         "Lists typical symptoms and preventive measures",
         "Includes solution approaches that have worked",
         "Speeds diagnosis and prevents redundant investigation"],
        "Knowledge base documents: common root causes in your systems, typical symptoms, preventive measures, and solution approaches. An accessible knowledge base speeds diagnosis and prevents redundant investigation."
    ),
    (
        "Explain defect change impact analysis.",
        ["Assessing how planned code changes might introduce defects",
         "Performed before deploying changes",
         "Identifies existing functionality that might be affected",
         "Determines what testing is needed as a result",
         "Proactive analysis prevents unexpected defects"],
        "Before deploying changes, impact analysis considers: what existing functionality might be affected, what testing is needed, and what defects might result. Proactive analysis prevents unexpected defects."
    ),
    (
        "What is a defect quality assurance score?",
        ["Metric measuring overall defect management quality",
         "Combines defect data accuracy, timely resolution, SLA compliance",
         "Includes proper documentation score",
         "Indicates how well defect process is being executed",
         "Identifies specific improvement areas"],
        "QA score combines: defect data accuracy, timely resolution, SLA compliance, and proper documentation. Score indicates how well the defect process is being executed and identifies improvement areas."
    ),
    (
        "Define defect management ROI.",
        ["Calculating business value of the defect management investment",
         "Cost: defect management process and tooling",
         "Benefits: prevented escapes, reduced support costs",
         "Improved customer satisfaction and faster time to market",
         "Positive ROI justifies investment in process and tools"],
        "ROI considers: cost of defect management process vs. value from: prevented escapes, reduced support costs, improved customer satisfaction, and faster time to market. Positive ROI justifies investment."
    ),
    (
        "What is a defect champion/mentor?",
        ["Experienced person leading and embedding defect process",
         "Expert in defect management best practices",
         "Mentors others and maintains process discipline",
         "Coaches teams on effective defect reporting and triage",
         "Drives continuous improvement and embeds process in culture"],
        "A defect champion is an expert in defect management who mentors others, maintains discipline in process, coaches teams on best practices, and drives continuous improvement. Champions embed the process into organisational culture."
    ),
    (
        "Explain defect management in Agile/DevOps.",
        ["Lightweight process for rapid iteration environments",
         "Minimal documentation to maintain speed",
         "Rapid triage — same day or same sprint",
         "Continuous testing automatically feeds defects",
         "Traditional detailed processes don't fit rapid environments"],
        "In Agile/DevOps, defect management must be lightweight: minimal documentation, rapid triage, continuous testing feeds defects, rapid deployment of fixes. Traditional detailed defect processes don't fit rapid environments."
    ),
    (
        "What is a defect management SOP?",
        ["Documented step-by-step standard operating procedure",
         "How to report a defect correctly",
         "Severity and priority assessment process",
         "Assignment, resolution, and verification workflow",
         "Ensures consistency and trains new team members"],
        "SOP documents: how to report a defect, severity/priority assessment process, assignment process, resolution workflow, verification process, and closure criteria. SOP ensures consistency and trains new team members."
    ),
    (
        "Define defect management culture.",
        ["Organisational mindset toward defects and quality",
         "Positive culture: defects are learning opportunities",
         "Quality is everyone's responsibility, not just QA",
         "Transparency about issues is encouraged",
         "Poor culture hides defects and blames individuals"],
        "Positive culture views: defects as learning opportunities (not failures), quality as everyone's responsibility, transparency about issues, and continuous improvement. Poor culture hides defects, blames individuals, resists process."
    ),
    (
        "What is the Defect Manager's role in product quality?",
        ["Drives quality through systematic defect management",
         "Ensures systematic defect detection and prioritisation",
         "Facilitates timely resolution of important issues",
         "Prevents recurrence through root cause analysis",
         "Analyses trends to drive continuous improvement"],
        "The Defect Manager drives quality by: ensuring systematic defect detection, prioritising important issues, facilitating resolution, preventing recurrence through RCA, and analysing trends to drive continuous improvement."
    ),
    (
        "What is the strategic importance of Defect Manager in organisation?",
        ["Critical to product quality and release readiness",
         "Prevents production defects through rigorous process",
         "Provides visibility into quality status for leadership",
         "Enables informed decisions about release readiness",
         "Drives quality investment and process improvement"],
        "The Defect Manager is critical to quality by ensuring effective defect identification, resolution, and prevention; providing visibility into quality status; preventing production defects through rigorous process; and enabling leadership to make informed decisions about release readiness and quality investments."
    ),
]


# ── 100 SAP Tester Q&A ────────────────────────────────────────────────────
SAP_TESTER_QA = [
    (
        "What is SAP?",
        ["Enterprise Resource Planning (ERP) system",
         "Integrates business processes across all functions",
         "Covers finance, HR, supply chain, manufacturing, sales",
         "Provides centralised data management",
         "Delivers real-time business insights"],
        "SAP (Systems, Applications, and Products in Data Processing) is a comprehensive Enterprise Resource Planning (ERP) software that integrates business processes across finance, human resources, supply chain, manufacturing, sales, and other functions. SAP provides centralised data management and real-time business insights across the organisation."
    ),
    (
        "Define SAP modules.",
        ["Functional components covering specific business areas",
         "FI (Financial Accounting), CO (Controlling)",
         "MM (Materials Management), SD (Sales & Distribution)",
         "PP (Production Planning), HCM (Human Capital Management)",
         "PS (Project Systems), QM (Quality Management)"],
        "SAP modules are functional components: Financial Accounting (FI), Controlling (CO), Materials Management (MM), Production Planning (PP), Sales & Distribution (SD), Human Capital Management (HCM), Project Systems (PS), and Quality Management (QM). Each module manages a specific business area."
    ),
    (
        "What is the role of an SAP Tester?",
        ["Validates SAP system meets business requirements",
         "Tests configuration, functionality, and integration",
         "Covers functional, integration, system, and UAT",
         "Verifies performance meets standards",
         "Ensures system is production-ready"],
        "An SAP Tester validates that the SAP system meets business requirements, is properly configured, functions correctly, integrates with other systems, and meets performance standards. Testing includes functional testing, integration testing, system testing, and UAT."
    ),
    (
        "Explain the SAP system architecture.",
        ["Multi-tier technical structure",
         "Presentation layer: user interface",
         "Application layer: business logic processing",
         "Database layer: data storage",
         "Communication between layers enables data flow"],
        "SAP architecture includes: Presentation layer (user interface), Application layer (business logic), and Database layer (data storage). Communication between layers enables data flow and processing. Understanding architecture helps testers understand system behaviour."
    ),
    (
        "What is the SAP GUI (Graphical User Interface)?",
        ["Client application for accessing SAP system",
         "Displays screens and processes user input",
         "Communicates with SAP application server",
         "Testers must be proficient in SAP GUI navigation",
         "Modern SAP also offers Fiori web-based interface"],
        "SAP GUI is the client application through which users access SAP. It displays screens, processes input, communicates with SAP server. Testers must be proficient in SAP GUI navigation and transaction execution. Modern SAP also offers web-based interfaces (Fiori)."
    ),
    (
        "Define SAP transaction code.",
        ["Short code to access a specific SAP function",
         "Examples: FB01 (post invoice), VA01 (create sales order)",
         "MB01 (create material), ME21N (create purchase order)",
         "Enables efficient navigation in SAP",
         "Knowing key t-codes for tested area is essential"],
        "Transaction codes (t-codes) are shortcuts to access SAP functions: FB01 (post invoice), MB01 (create material), VA01 (create sales order). Testers use t-codes to navigate efficiently. Knowing key t-codes for the tested area is essential."
    ),
    (
        "What is SAP Customising (IMG)?",
        ["Configuration of SAP to meet business requirements",
         "IMG = Implementation Guide",
         "Includes organisational structures and process flows",
         "Functional consultants perform customising",
         "Testers validate customising is correctly applied"],
        "Customising (Implementation Guide or IMG) is configuration of the SAP system to match business needs. It includes: setting up organisational structures, defining process flows, configuring master data, and establishing rules. Functional consultants perform customising; testers validate it works."
    ),
    (
        "Explain SAP master data.",
        ["Foundational reference data in SAP",
         "Includes materials, vendors, customers, employees",
         "Also cost centres and general ledger accounts",
         "Foundation for all transactions",
         "Incorrect master data causes downstream errors"],
        "Master data includes: materials (products), vendors, customers, employees, cost centres, and general ledger accounts. Master data is the foundation for transactions; incorrect master data causes downstream errors. Testers should validate master data setup."
    ),
    (
        "What is SAP transactional data?",
        ["Records of business activities in SAP",
         "Sales orders, purchase orders, invoices, payments",
         "Created from master data",
         "Testing validates correct creation and processing",
         "Results must be accurate and complete"],
        "Transactional data represents business activities: sales orders, purchase orders, invoices, payments. Transactional data is created from master data. Testing validates transactions are created correctly, processed properly, and produce accurate results."
    ),
    (
        "Define SAP table.",
        ["Database structure storing SAP data",
         "MARA: materials master data",
         "VBAK: sales order header",
         "KNA1: customer master",
         "Testers use tables to verify data accuracy"],
        "SAP tables store data in the database: MARA (materials master), VBAK (sales order header), KNA1 (customer master). Understanding key tables helps testers verify data accuracy. Testers should know tables relevant to tested processes."
    ),
    (
        "What is a business process in SAP?",
        ["End-to-end workflow spanning multiple transactions",
         "Procure-to-Pay: purchase order through payment",
         "Order-to-Cash: order through invoice collection",
         "Plan-to-Produce: planning through manufacturing",
         "Testers validate complete process accuracy"],
        "Business processes are end-to-end workflows: Procure-to-Pay (purchase order through payment), Order-to-Cash (order through invoice), Plan-to-Produce (planning through manufacturing). Testers validate entire processes work correctly."
    ),
    (
        "Explain the Procure-to-Pay (P2P) process.",
        ["Purchase Requisition → Purchase Order → Goods Receipt",
         "Then: Invoice → Payment",
         "Each step involves specific transactions",
         "Produces accurate financial records",
         "P2P testing ensures all steps integrate correctly"],
        "P2P process: Purchase Requisition → Purchase Order → Goods Receipt → Invoice → Payment. Each step involves specific transactions and system updates. P2P testing ensures all steps integrate correctly and produce accurate financial records."
    ),
    (
        "What is the Order-to-Cash (O2C) process?",
        ["Sales Order → Delivery → Billing → Payment",
         "Each step involves transaction processing",
         "Inventory updates and financial recording",
         "Validates customer orders process correctly",
         "Covers entire flow from sale through cash collection"],
        "O2C process: Sales Order → Delivery → Billing → Payment. Each step involves transaction processing, inventory updates, and financial recording. O2C testing validates customer orders process correctly from sale through collection."
    ),
    (
        "Define SAP testing scope.",
        ["What aspects of SAP are tested",
         "Functionality, integration, data accuracy",
         "Performance and security",
         "Scope depends on implementation phase",
         "Must align with business requirements"],
        "Testing scope includes: functionality (features work as specified), integration (modules exchange data correctly), data accuracy (correct calculations and updates), performance (system responds timely), and security (access controls work). Scope depends on implementation phase."
    ),
    (
        "What is functional testing in SAP?",
        ["Verifies SAP functionality meets requirements",
         "Tests transaction execution: purchases, sales, payments",
         "Validates business logic and calculations",
         "Covers master data maintenance and reporting",
         "Testers execute test cases against configured SAP"],
        "Functional testing validates: transaction execution (purchases, sales, payments work), business logic (calculations correct), master data maintenance, reporting, and process flows. Testers execute test cases against the configured SAP system."
    ),
    (
        "Explain integration testing in SAP.",
        ["Tests data flow between SAP modules",
         "Purchase order in MM triggers PO in FI",
         "Delivery updates inventory and creates invoice",
         "No data loss occurs between modules",
         "Critical for end-to-end process validation"],
        "Integration testing verifies: modules exchange data correctly (purchase order in MM triggers PO in FI), interdependencies work (delivery updates inventory and creates invoice), no data loss occurs between modules. Critical for end-to-end process validation."
    ),
    (
        "What is SAP configuration testing?",
        ["Verifying configuration matches business requirements",
         "Confirms organisational structures are set up correctly",
         "Process variants configured as specified",
         "Business rules are properly implemented",
         "Configuration errors cause systematic process failures"],
        "Configuration testing confirms: organisational structures are set up correctly, process variants are configured as specified, business rules are implemented, master data is valid. Configuration errors cause systematic process failures."
    ),
    (
        "Define SAP interface testing.",
        ["Testing data exchange with external systems",
         "SAP integrates with EDI, supplier ERPs, data warehouses",
         "Verifies data transfers correctly to/from SAP",
         "Formats and frequencies must match requirements",
         "No data loss allowed during transfer"],
        "SAP often integrates with external systems (EDI, ERP of suppliers, data warehouses). Interface testing verifies: data transfers correctly from/to SAP, formats are correct, no loss occurs, frequencies match requirements."
    ),
    (
        "What is SAP performance testing?",
        ["Evaluating system speed under expected load",
         "Measures transaction response times",
         "Tests batch job execution time",
         "Validates report generation speed",
         "Evaluates behaviour under concurrent users"],
        "Performance testing measures: transaction response time (purchase order creation should complete in X seconds), batch job execution time, report generation speed, and system behaviour under concurrent users. Performance issues affect user productivity."
    ),
    (
        "Explain SAP data migration testing.",
        ["Validating accuracy of data moved from legacy system",
         "Verifies all legacy data is converted accurately",
         "Completeness: no data lost",
         "Referential integrity: master/transaction data consistency",
         "Incorrect migration causes major post-go-live issues"],
        "Data migration testing verifies: all legacy data is converted accurately, completeness (no data lost), referential integrity (master/transaction data consistency). Incorrect migration causes major post-go-live issues."
    ),
    (
        "What is SAP regression testing?",
        ["Re-testing after changes to ensure no breakage",
         "Verifies configuration changes don't break functionality",
         "Ensures fixes don't introduce new problems",
         "Patches must not impact previously tested areas",
         "Critical given SAP's complexity and interdependencies"],
        "Regression testing in SAP verifies: configuration changes don't break existing functionality, fixes don't introduce new problems, patches don't impact tested areas. Critical given the complexity of SAP."
    ),
    (
        "Define SAP test case.",
        ["Detailed script for testing specific functionality",
         "Documents what to test (transaction)",
         "Includes preconditions and detailed steps",
         "States expected results and acceptance criteria",
         "Enables consistent testing and sign-off evidence"],
        "An SAP test case documents: what to test (transaction), preconditions, detailed steps, expected results, acceptance criteria. Well-written cases enable consistent testing and provide evidence for sign-off."
    ),
    (
        "What is SAP test data?",
        ["Information needed to execute tests",
         "Includes sample materials, customers, vendors",
         "Also purchase orders and sales orders",
         "Should be realistic and representative of production",
         "Must be independent across test cases"],
        "SAP test data includes: sample materials, customers, vendors, purchase orders, sales orders needed to execute tests. Test data should be realistic, representative of production scenarios, and independent across test cases."
    ),
    (
        "Explain SAP sandbox/test environment.",
        ["Non-production system for testing",
         "Copy of production SAP configuration",
         "Testing occurs without affecting live data",
         "Allows experimentation and failure without consequences",
         "Multiple sandboxes support different testing phases"],
        "Sandbox is a copy of production SAP configuration where testing occurs without affecting live data. Testing in sandbox allows experimentation, failure without consequences, and validation before production. Multiple sandboxes support different testing phases."
    ),
    (
        "What is SAP golden configuration?",
        ["Approved configuration used as testing reference",
         "Represents final production configuration",
         "Serves as reference for validation testing",
         "Enables rapid environment refresh",
         "Provides reproducible testing baseline"],
        "Golden configuration is the approved, tested SAP setup representing the final production configuration. It serves as a reference for validation testing, enables rapid environment refresh, and provides a reproducible baseline."
    ),
    (
        "Define SAP GRC (Governance, Risk, Compliance).",
        ["Controls and audit features in SAP",
         "Ensures segregation of duties",
         "Audit trails record all changes",
         "Access controls restrict data visibility",
         "GRC testing validates controls are effective"],
        "GRC functionality ensures: segregation of duties (no single person approving and paying), audit trails record all changes, access controls restrict data visibility. GRC testing validates controls are effective."
    ),
    (
        "What is segregation of duties (SoD) in SAP?",
        ["Preventing fraud through separation of conflicting duties",
         "Creating purchase order AND approving payment is a conflict",
         "Roles restrict users from incompatible transactions",
         "Testing validates role assignments enforce SoD",
         "SoD failures create fraud risk"],
        "SoD ensures no single person performs conflicting duties: creating purchase orders AND approving payment would be a conflict. Testing validates roles restrict users from incompatible transactions, preventing fraud."
    ),
    (
        "Explain SAP authorisation testing.",
        ["Verifying user access controls work correctly",
         "Users can only access assigned functions",
         "Role assignments are appropriate and complete",
         "Sensitive data is restricted to authorised users",
         "No unauthorised access is possible"],
        "Authorisation testing confirms: users can access only assigned functions, role assignments are appropriate, sensitive data is restricted, and no unauthorised access is possible. Authorisation failures create security risks."
    ),
    (
        "What is SAP report testing?",
        ["Validating SAP reports produce accurate data",
         "Reports extract correct data",
         "Calculations are accurate",
         "Formatting is appropriate for business use",
         "Performance is acceptable for users"],
        "Report testing verifies: reports extract correct data, calculations are accurate, formatting is appropriate, performance is acceptable. Reports are critical for business decision-making; accuracy is essential."
    ),
    (
        "Define SAP query/report variants.",
        ["Pre-configured report filters and outputs",
         "Saved report configurations showing specific data subsets",
         "Example: sales report for Region A, Q1",
         "Testing verifies variants filter and display correctly",
         "Variants produce expected output consistently"],
        "Variants are saved report configurations showing a specific data subset (e.g., sales report for Region A, Q1). Testing variants ensures they filter/display correctly and produce expected outputs."
    ),
    (
        "What is SAP Fiori?",
        ["Modern, user-friendly interface for SAP",
         "Simplifies complex SAP transactions",
         "Web-based and mobile-responsive",
         "Provides consumer-grade user experience",
         "Testers must be proficient in both GUI and Fiori"],
        "Fiori is SAP's modern, user-friendly interface providing a simplified user experience. Fiori apps simplify complex SAP transactions. Testers should be proficient in both SAP GUI and Fiori testing."
    ),
    (
        "Explain SAP ABAP.",
        ["SAP's programming language for customisation",
         "Used for custom reports, interfaces, and enhancements",
         "Testers need basic ABAP knowledge",
         "ABAP testing validates custom code works correctly",
         "Custom code often requires targeted unit testing"],
        "ABAP is SAP's programming language for custom reports, interfaces, and enhancements. Testers need basic ABAP knowledge to understand custom code being tested. ABAP testing validates custom code works correctly."
    ),
    (
        "What is SAP RFC (Remote Function Call)?",
        ["Mechanism for communication between SAP and external systems",
         "Enables communication between SAP systems",
         "RFC testing verifies calls execute successfully",
         "Data transfers correctly via RFC",
         "Error handling must work for failed calls"],
        "RFC enables communication between SAP and external systems or between SAP systems. RFC testing verifies: calls execute successfully, data transfers correctly, error handling works."
    ),
    (
        "Define SAP batch processing.",
        ["Scheduled background jobs in SAP",
         "Runs payroll processing, month-end closings, data cleanup",
         "Batch testing verifies jobs execute on schedule",
         "Jobs must complete successfully within time windows",
         "Output must be accurate and complete"],
        "Batch jobs run scheduled activities (payroll processing, month-end closings, data cleanup). Batch testing verifies: jobs execute on schedule, complete successfully, produce accurate output."
    ),
    (
        "What is SAP month-end close process?",
        ["Financial closing activities at month end",
         "Includes accruals, reconciliations, and closing entries",
         "Executed in FI module",
         "Testing validates all close steps complete correctly",
         "Reports must reflect correctly closed results"],
        "Month-end close includes: accruals, reconciliations, consolidations, and closing entries in FI module. Testing validates: all close steps complete correctly, balances are correct, reports reflect closed results."
    ),
    (
        "Explain SAP year-end close process.",
        ["Comprehensive financial closing at year end",
         "All month-end activities plus audits and adjustments",
         "Includes financial statement preparation",
         "Year-end testing is extensive and critical",
         "All financial data must be accurate before close"],
        "Year-end close is comprehensive: all month-end activities plus audits, adjustments, financial statement preparation. Year-end testing validates: accuracy of all financial data, completion of all required close steps."
    ),
    (
        "What is SAP Controlling (CO) module?",
        ["Cost management and internal reporting module",
         "Tracks costs and allocates to cost centres",
         "Provides internal profitability reporting",
         "CO testing validates cost allocations are correct",
         "Internal reports must be accurate"],
        "CO module tracks costs, allocates costs to cost centres, provides internal profitability reporting. CO testing validates: cost allocations are correct, internal reports are accurate, cost centre tracking is complete."
    ),
    (
        "Define SAP cost allocation.",
        ["Distribution of indirect costs to cost centres",
         "Distributes overhead based on defined allocation rules",
         "Testing validates allocation basis is correct",
         "Amounts must be accurate after allocation",
         "Incorrect allocation distorts profitability reporting"],
        "Cost allocation distributes overhead costs based on defined rules. Testing validates: allocation basis is correct, amounts are accurate, allocation completes successfully. Incorrect allocation distorts profitability."
    ),
    (
        "What is SAP Materials Management (MM)?",
        ["Inventory and procurement functions module",
         "Manages purchasing, receiving, inventory, materials planning",
         "MM testing validates purchase orders created correctly",
         "Goods receipts must update inventory accurately",
         "Procurement cycles work end-to-end"],
        "MM module manages: purchasing, receiving, inventory, materials planning. MM testing validates: purchase orders are created correctly, goods receipts update inventory, inventory accuracy, procurement cycles work end-to-end."
    ),
    (
        "Explain SAP Sales & Distribution (SD).",
        ["Sales order and delivery functions module",
         "Manages customer orders, deliveries, billing, sales reporting",
         "SD testing validates orders are created correctly",
         "Deliveries update inventory and trigger billing",
         "Invoices must be accurate and complete"],
        "SD module manages: customer orders, deliveries, billing, sales reporting. SD testing validates: orders are created correctly, deliveries update inventory and trigger billing, invoices are accurate, sales data is complete."
    ),
    (
        "What is SAP Production Planning (PP)?",
        ["Manufacturing planning and execution module",
         "Manages demand planning, production orders, scheduling",
         "PP testing validates production orders created correctly",
         "Materials must be allocated properly",
         "Completions update inventory accurately"],
        "PP module manages: demand planning, production orders, work-in-progress, production scheduling. PP testing validates: production orders are created correctly, materials are allocated, completions update inventory accurately."
    ),
    (
        "Define SAP Human Capital Management (HCM).",
        ["Employee and payroll functions module",
         "Manages employee records, payroll, benefits, attendance",
         "HCM testing validates employee data is correct",
         "Payroll calculations must be accurate",
         "Deductions applied correctly and reports complete"],
        "HCM manages: employee records, payroll, benefits, attendance. HCM testing validates: employee data is correct, payroll calculations are accurate, deductions are applied correctly, reports are complete."
    ),
    (
        "What is SAP project testing approach?",
        ["Comprehensive testing strategy for SAP implementations",
         "Unit: individual transactions",
         "Integration: cross-module data flow",
         "System: complete environment validation",
         "UAT: business validation, then cutover"],
        "SAP project testing includes: Unit (individual transactions), Integration (cross-module), System (complete environment), UAT (business validation), cutover (production readiness). Each phase validates different aspects."
    ),
    (
        "Explain SAP test strategy in ASAP.",
        ["Testing embedded throughout ASAP implementation methodology",
         "Blueprint phase: define test requirements",
         "Realization: test configuration as built",
         "Final Preparation: regression and UAT",
         "Continuous testing reduces late-cycle defects"],
        "ASAP (Accelerated SAP) integrates testing throughout: Blueprint phase (test requirements), Realization (test configuration), Final Preparation (regression and UAT). Continuous testing throughout implementation reduces late-cycle defects."
    ),
    (
        "What is SAP process variant testing?",
        ["Testing alternate configurations of the same process",
         "Variants represent different business configurations",
         "Regional variations or different operating companies",
         "Each variant must be individually tested",
         "Ensures all configured variants work correctly"],
        "Process variants are different ways to configure business processes in SAP. If the organisation supports multiple variants (regional variations, different operating companies), each variant must be tested to ensure it works."
    ),
    (
        "Define SAP localisation testing.",
        ["Testing for specific country or regional requirements",
         "Ensures local tax requirements are met",
         "Local language support works correctly",
         "Local regulatory features function as required",
         "Each deployment geography must be validated"],
        "Localisation testing ensures: local tax requirements are met, local language support works, local regulatory features function. SAP's multi-country capability requires localisation validation for each deployment geography."
    ),
    (
        "What is SAP upgrade testing?",
        ["Testing after SAP version upgrade",
         "Verifies existing functionality works in new version",
         "Customisations must be compatible",
         "Data integrity maintained after upgrade",
         "Performance must be acceptable in new version"],
        "Upgrade testing verifies: existing functionality works in new version, customisations are compatible, data integrity is maintained, performance is acceptable in new version. Upgrade testing is critical before production cutover."
    ),
    (
        "Explain SAP patch testing.",
        ["Validation after SAP support packs or patches applied",
         "Patches fix bugs and provide enhancements",
         "Verifies fixes work correctly",
         "No new issues introduced by patch",
         "Patches require regression testing"],
        "Patches fix bugs and provide enhancements. Patch testing verifies: fixes work, no new issues introduced, no breakage of other functionality, system stability. Patches require regression testing."
    ),
    (
        "What is SAP golden record testing?",
        ["Testing with representative production-like data",
         "Uses realistic scenarios, quantities, and data",
         "Provides confidence system handles real data",
         "Performed just before go-live",
         "Closest to real-world production conditions"],
        "Golden record testing uses realistic scenarios, quantities, and data representative of actual production. Provides confidence that the system will handle real data correctly. Performed just before go-live."
    ),
    (
        "Define SAP stress testing.",
        ["Testing system behaviour under extreme load",
         "Pushes system to its limits",
         "High transaction volume and large batch jobs",
         "Many concurrent users simulated",
         "Identifies breaking points and resource bottlenecks"],
        "Stress testing pushes the system to limits: high transaction volume, large batch jobs, many concurrent users. Identifies breaking points, resource bottlenecks, and system stability issues under peak loads."
    ),
    (
        "What is SAP cutover testing?",
        ["Final validation before production go-live",
         "Verifies all go-live preparations are complete",
         "Data migration must be accurate",
         "Interfaces must be operational",
         "Acts as gate for go-live decision"],
        "Cutover testing verifies: all go-live preparations are complete, system is production-ready, data migration is accurate, interfaces are operational, disaster recovery is configured. Cutover testing is the gate for go-live."
    ),
    (
        "Explain SAP hypercare period.",
        ["Intensive support period after go-live",
         "Extended support team available",
         "Testers involved in rapid defect verification",
         "Workaround testing during critical period",
         "Issues require immediate resolution"],
        "Hypercare is the post-go-live period with extended support team. Testers may be involved in rapid defect verification and workaround testing during this critical period. Issues discovered post-go-live require immediate resolution."
    ),
    (
        "What is SAP end-to-end (E2E) testing?",
        ["Testing complete business processes from start to finish",
         "Sales order through delivery and payment",
         "Purchase requisition through payment",
         "All integrated steps must work correctly",
         "Results must be accurate at every step"],
        "E2E testing traces processes from beginning to end: sales order through delivery and payment, purchase requisition through payment. Validates all integrated steps work correctly and produce accurate results."
    ),
    (
        "Define SAP smoke testing.",
        ["Quick verification after a build or deployment",
         "After configuration update or patch",
         "Quick tests verify system basics still work",
         "Identifies major issues before detailed testing",
         "Fast and shallow — not exhaustive"],
        "Smoke testing in SAP context: after configuration update or patch, quick tests verify system basics still work (can log in, basic transactions execute). Identifies major issues before detailed testing begins."
    ),
    (
        "What is SAP master data testing?",
        ["Validating reference data setup",
         "Materials configured correctly",
         "Customers and vendors set up properly",
         "Accounts exist in chart of accounts",
         "Master data errors cascade through transactions"],
        "Master data testing ensures: materials are configured correctly, customers are set up properly, vendors have required data, accounts are in the chart of accounts. Master data errors cascade through transactions."
    ),
    (
        "Explain SAP document testing.",
        ["Validating creation and processing of SAP documents",
         "Covers purchase orders, sales orders, invoices, payments",
         "Documents created with required data",
         "Process through system accurately",
         "Produce expected financial impacts"],
        "Document testing covers: purchase orders, sales orders, invoices, payment documents. Tests verify: documents are created correctly, contain required data, process through the system accurately, and produce expected financial impacts."
    ),
    (
        "What is SAP posting testing?",
        ["Verifying financial entries in the general ledger",
         "Correct amounts to correct G/L accounts",
         "Accounts are balanced after posting",
         "Financial reports reflect postings accurately",
         "Incorrect posting creates financial inaccuracy"],
        "Posting testing validates: transactions post correct amounts to correct G/L accounts, accounts are balanced, financial reports reflect postings accurately. Incorrect posting creates financial inaccuracy."
    ),
    (
        "Define SAP clearing testing.",
        ["Validating invoice and payment matching",
         "Invoices and payments match correctly",
         "Clearing processes work as configured",
         "Accounts receivable/payable aging is accurate",
         "Open items are tracked properly"],
        "Clearing testing ensures: invoices and payments match correctly, clearing processes work, accounts receivable/payable aging is accurate, open items are tracked properly."
    ),
    (
        "What is SAP reconciliation testing?",
        ["Verifying account balances match between systems",
         "SAP general ledger matches subledgers",
         "Inventory values reconcile",
         "Customer and vendor balances reconcile",
         "Reconciliation catches data integrity issues"],
        "Reconciliation testing validates: SAP general ledger balances match subledgers, inventory values match, customer/vendor balances reconcile. Reconciliation catches data integrity issues."
    ),
    (
        "Explain SAP balance testing.",
        ["Ensuring accounts balance correctly",
         "Debits equal credits",
         "Balance sheet balances",
         "Assets equal liabilities plus equity",
         "Balance errors reveal data quality or configuration issues"],
        "Balance testing verifies: debits equal credits, balance sheet balances, assets equal liabilities plus equity. Balance errors reveal data quality issues or configuration problems."
    ),
    (
        "What is SAP FIFO/LIFO testing in MM?",
        ["Testing inventory valuation methods",
         "SAP supports FIFO, LIFO, and standard cost",
         "Selected method must be applied correctly",
         "Inventory valued accurately using chosen method",
         "Variance reporting must be correct"],
        "SAP supports different inventory valuation (FIFO, LIFO, standard cost). Testing validates: selected method is applied correctly, inventory is valued accurately, variance reporting is correct."
    ),
    (
        "Define SAP reorder point testing.",
        ["Testing automated purchase requisition creation",
         "Reorder point triggers requisitions when inventory is low",
         "Reorder points must be set correctly",
         "Requisitions generate when thresholds are reached",
         "Quantities must be appropriate"],
        "Reorder point feature automatically creates purchase requisitions when inventory falls below threshold. Testing verifies: reorder points are set correctly, requisitions generate when thresholds are reached, quantities are appropriate."
    ),
    (
        "What is SAP forecast testing?",
        ["Validating demand planning functionality",
         "Demand calculated correctly",
         "Seasonal factors applied properly",
         "Forecast accuracy is reasonable",
         "Production planning uses forecast appropriately"],
        "Forecast testing ensures: demand is calculated correctly, seasonal factors are applied, forecast accuracy is reasonable, production planning uses forecast appropriately."
    ),
    (
        "Explain SAP bill of materials (BOM) testing.",
        ["Validating product component structure",
         "Component lists are accurate",
         "Quantities are correct",
         "Alternate BOM variants work",
         "BOM explosions in production orders are correct"],
        "BOM testing verifies: component lists are accurate, quantities are correct, alternate BOM variants work, BOM explosions in production orders are correct. Incorrect BOMs cause production errors."
    ),
    (
        "What is SAP routing testing?",
        ["Validating production sequence and operations",
         "Work centres are assigned correctly",
         "Production sequences are logical",
         "Time estimates are reasonable",
         "Production orders follow defined routings"],
        "Routing defines production steps for manufacturing. Testing verifies: work centres are assigned correctly, sequences are logical, time estimates are reasonable, production orders follow routings."
    ),
    (
        "Define SAP quality management (QM) testing.",
        ["Testing quality control functionality",
         "Inspection lots are created correctly",
         "Quality checks execute correctly",
         "Defects are tracked in QM",
         "Material quarantined when quality fails"],
        "QM testing validates: inspection lots are created, quality checks execute correctly, defects are tracked, material is quarantined if needed. QM testing ensures quality processes function."
    ),
    (
        "What is SAP batch management testing?",
        ["Testing lot traceability functionality",
         "Batches created and tracked correctly",
         "Expiration dates monitored",
         "Lot genealogy tracked throughout lifecycle",
         "Critical for industries with batch tracking requirements"],
        "Batch testing validates: batches are created and tracked, expiration dates are monitored, lot genealogy is tracked, batch-level quality results are recorded. Critical for industries with batch tracking."
    ),
    (
        "Explain SAP serial number testing.",
        ["Testing individual item tracking",
         "Serial numbers generated and tracked",
         "Equipment tracked throughout lifecycle",
         "Warranty information associated with serial",
         "Recalls can target specific serial numbers"],
        "Serial number testing verifies: serial numbers are generated and tracked, equipment is tracked throughout lifecycle, warranty information is associated, recalls can target serial numbers."
    ),
    (
        "What is SAP tax testing?",
        ["Validating tax calculations and reporting",
         "Tax rates applied correctly",
         "Tax amounts calculated accurately",
         "Tax reports include all taxable transactions",
         "Tax compliance requirements are met"],
        "Tax testing ensures: tax rates are applied correctly, tax amounts are calculated accurately, tax reports include all taxable transactions, tax compliance is met. Tax accuracy is critical and often audited."
    ),
    (
        "Define SAP profit centre testing.",
        ["Testing internal profit and loss tracking",
         "Transactions assigned to correct profit centres",
         "Profit centre accounting produces accurate results",
         "Management reporting shows correct profitability",
         "Profit centre data must reconcile"],
        "Profit centres track P&L by business unit. Testing validates: transactions are assigned to correct profit centres, profit centre accounting produces accurate results, management reporting shows correct profitability by centre."
    ),
    (
        "What is SAP cost element testing?",
        ["Validating expense categorisation",
         "Expenses charged to correct cost elements",
         "Cost element reporting is accurate",
         "Cost analysis by element is meaningful",
         "Provides basis for management reporting"],
        "Cost elements categorise costs (salaries, supplies, utilities). Testing verifies: expenses are charged to correct cost elements, cost element reporting is accurate, cost analysis by element is meaningful."
    ),
    (
        "Explain SAP ledger group testing.",
        ["Testing multiple accounting ledgers",
         "Transactions recorded in all configured ledgers",
         "Ledger reconciliation is correct",
         "Parallel reporting produces expected differences",
         "Enables multiple accounting standards simultaneously"],
        "Ledger groups allow multiple parallel accounting methods (parallel reporting). Testing validates: transactions are recorded in all configured ledgers, ledger reconciliation is correct, parallel reporting produces expected differences."
    ),
    (
        "What is SAP dunning testing?",
        ["Testing collection/reminder functionality",
         "Dunning levels trigger at correct overdue days",
         "Dunning letters generate correctly",
         "Dunning status is tracked accurately",
         "Automates collection reminders for overdue amounts"],
        "Dunning automates collection reminders for overdue amounts. Testing verifies: dunning levels trigger at correct overdue days, dunning letters generate correctly, dunning status is tracked."
    ),
    (
        "Define SAP rebate testing.",
        ["Testing volume discount calculations",
         "Discount rates applied correctly",
         "Rebate amounts calculated accurately",
         "Rebate conditions evaluated properly",
         "Rebate accruals booked correctly"],
        "Rebate testing validates: discount rates are applied correctly, rebate amounts are calculated accurately, rebate conditions are evaluated properly, rebate accruals are booked."
    ),
    (
        "What is SAP intercompany testing?",
        ["Testing transactions between corporate entities",
         "Sales between company entities recorded correctly",
         "Intercompany clearing works",
         "Consolidation adjustments identified",
         "No balances exist between entities in consolidated view"],
        "Intercompany testing verifies: sales between company entities are recorded, intercompany clearing works, consolidation adjustments are needed, no balances exist between entities (for consolidated reporting)."
    ),
    (
        "Explain SAP exchange rate testing.",
        ["Testing multi-currency functionality",
         "Transactions posted in transaction and reporting currency",
         "Exchange gains/losses calculated correctly",
         "Year-end revaluations recorded correctly",
         "Currency conversion accurate at all stages"],
        "Exchange rate testing validates: transactions are posted in both transaction currency and reporting currency, exchange gains/losses are calculated, year-end revaluations are recorded correctly."
    ),
    (
        "What is SAP variance analysis testing?",
        ["Testing standard cost variance calculations",
         "Actual costs compared to standard costs",
         "Material price variance calculated correctly",
         "Labour efficiency variance recorded",
         "Variances reconcile to P&L"],
        "Variance testing ensures: actual costs are compared to standard costs, material price variance is calculated, labour efficiency variance is recorded, variances reconcile to P&L."
    ),
    (
        "Define SAP consolidation testing.",
        ["Testing multi-entity financial consolidation",
         "Intercompany transactions are eliminated",
         "Minority interests are recorded",
         "Consolidated financial statements are accurate",
         "Elimination entries are complete"],
        "Consolidation testing validates: intercompany transactions are eliminated, minority interests are recorded, consolidated financial statements are accurate, elimination entries are complete."
    ),
    (
        "What is SAP working capital testing?",
        ["Testing cash flow and liquidity management",
         "Cash flow projections are accurate",
         "Receivables and payables are managed correctly",
         "Liquidity reporting shows cash position",
         "Working capital metrics calculated correctly"],
        "Working capital testing verifies: cash flow projections are accurate, receivables and payables are managed, liquidity reporting shows cash position, working capital metrics are calculated correctly."
    ),
    (
        "Explain SAP drill-down testing.",
        ["Testing navigation from summary to detail",
         "Users navigate from summary reports to detail",
         "Drill-down shows accurate supporting detail",
         "Data at each level reconciles",
         "Drill-down is essential for business analysis"],
        "Drill-down testing validates: users can navigate from summary reports to detailed data, drill-down shows accurate supporting detail, data at each level reconciles. Drill-down capability is essential for analysis."
    ),
    (
        "What is SAP currency conversion testing?",
        ["Testing translation of transactions to reporting currency",
         "Exchange rates applied correctly",
         "Gain/loss is calculated accurately",
         "Summary and detail reconcile in both currencies",
         "Year-end closing includes proper revaluations"],
        "Currency conversion testing ensures: exchange rates are applied correctly, gain/loss is calculated, summary and detail reconcile in both currencies, year-end closing includes proper revaluations."
    ),
    (
        "Define SAP ledger posting testing.",
        ["Ensuring transactions post to appropriate ledgers",
         "FI transactions post to General Ledger",
         "Subledger posting (A/R, A/P) is accurate",
         "Reconciliation between G/L and subledgers is complete",
         "Posting accuracy is fundamental to financial integrity"],
        "Testing validates: FI transactions post to General Ledger, subledger posting (A/R, A/P) is accurate, reconciliation between G/L and subledgers is complete. Posting accuracy is fundamental."
    ),
    (
        "What is SAP workflow testing?",
        ["Testing automated approval and routing",
         "Purchase requisitions routed for approval",
         "Invoices requiring authorisation processed",
         "Routing logic is correct",
         "Escalation works when approvals are delayed"],
        "SAP workflows automate approvals: purchase requisitions routed for approval, invoices requiring authorisation. Workflow testing validates: routing logic is correct, approvals function, escalation works."
    ),
    (
        "Explain SAP exception handling testing.",
        ["Testing error detection and correction",
         "System identifies exceptions: mismatches, errors",
         "Provides clear and helpful error messages",
         "Enables correction without losing data",
         "Prevents incorrect data from propagating"],
        "Exception testing verifies: system identifies exceptions (mismatches, errors), provides clear error messages, enables correction without losing data, prevents propagation of incorrect data."
    ),
    (
        "What is SAP SLA testing?",
        ["Testing system compliance to service level agreements",
         "Response times meet agreed SLAs",
         "Availability is within SLA targets",
         "Batch processing completes within SLA windows",
         "SLA non-compliance affects business operations"],
        "SLA testing validates: system response times meet SLA, availability is within SLA, batch processing completes within SLA windows, reporting meets SLA. SLA non-compliance affects business."
    ),
    (
        "Define SAP security testing.",
        ["Testing access controls and data protection",
         "Role-based access controls work correctly",
         "Passwords handled securely",
         "Encryption applied where required",
         "Audit logs capture all relevant activity"],
        "Security testing validates: role-based access controls work, passwords are handled securely, encryption is applied, audit logs capture activity, compliance requirements are met. Security breaches are catastrophic."
    ),
    (
        "What is SAP backup and recovery testing?",
        ["Validating data can be restored after issues",
         "Backups execute successfully",
         "Backup data is complete",
         "Recovery processes work as designed",
         "Recovered data is accurate and usable"],
        "Backup testing verifies: backups execute successfully, backup data is complete, recovery processes work, recovered data is accurate. Regular testing provides confidence in disaster recovery."
    ),
    (
        "Explain SAP volume testing.",
        ["Testing system with production-scale data",
         "Large product catalogues and millions of transactions",
         "Years of historical data included",
         "Validates HANA/SAP handles scale",
         "Volume testing reveals scalability issues"],
        "Volume testing uses realistic data quantities to assess performance and behaviour: large product catalogues, millions of transactions, years of history. Volume testing reveals scalability issues."
    ),
    (
        "What is SAP compatibility testing?",
        ["Testing with different platforms and software versions",
         "Validates SAP works with supported databases",
         "Tests across operating systems",
         "Covers browsers and third-party tools",
         "Ensures technical environment is fully compatible"],
        "Compatibility testing ensures: SAP works with supported databases, operating systems, browsers, third-party tools. Compatibility testing provides confidence in the technical environment."
    ),
    (
        "Define SAP user acceptance testing (UAT).",
        ["Business validation that system meets requirements",
         "Actual business users perform testing",
         "Confirms configured SAP meets their requirements",
         "Supports their business processes effectively",
         "UAT sign-off is the gate for go-live"],
        "UAT involves business users testing that configured SAP meets their requirements, supports their business processes, and is ready for use. UAT sign-off is the gate for go-live."
    ),
    (
        "What is SAP documentation testing?",
        ["Validating user guides and system documentation",
         "Help text is accurate and useful",
         "User guides are complete",
         "Transaction documentation is correct",
         "System messages are clear and helpful"],
        "Documentation testing verifies: help text is accurate, user guides are complete, transaction documentation is correct, system messages are helpful. Good documentation supports user adoption."
    ),
    (
        "Explain SAP training validation.",
        ["Testing that users can operate SAP after training",
         "Users understand SAP navigation",
         "Can execute key transactions independently",
         "Know how to handle common errors",
         "Understand business processes in SAP context"],
        "Training validation assesses: users understand SAP navigation, can execute key transactions, know how to handle errors, understand business processes in SAP. Training effectiveness is critical for success."
    ),
    (
        "What is SAP defect categorisation in SAP context?",
        ["Classifying SAP defects by type",
         "Configuration error: setup incorrect",
         "Data error: master or transactional data wrong",
         "Interface error: data transfer failed",
         "Custom code error: ABAP or script issue"],
        "SAP defects are categorised: Configuration error (setup incorrect), Data error (master/transactional data wrong), Interface error (data transfer failed), Custom code error (ABAP/script issue), Functional gap (SAP limitation). Category indicates where the fix occurs."
    ),
    (
        "Define SAP test environment management.",
        ["Managing configurations and data across test environments",
         "Separate environments: Dev, QA, UAT",
         "Configuration consistency across environments",
         "Data separation to prevent test interference",
         "Environment refresh capability"],
        "Environment management maintains: separate environments (Dev, QA, UAT), configuration consistency across environments, data separation to prevent test interference, environment refresh capability."
    ),
    (
        "What is SAP test automation in context?",
        ["Using tools to automate SAP testing",
         "Tools: UFT, ReadyAPI, Tricentis Tosca",
         "Automates regression testing",
         "Improves efficiency and test coverage",
         "Requires maintenance as SAP changes"],
        "SAP test automation tools (UFT, ReadyAPI, Tricentis) automate: regression testing, data validation, report verification. Automation improves efficiency but requires maintenance as SAP changes."
    ),
    (
        "Explain SAP testing best practices.",
        ["Thorough test planning from the outset",
         "Comprehensive requirements review",
         "Focus on business process testing",
         "Strong master data validation",
         "Rigorous regression testing and user involvement in UAT"],
        "Best practices include: thorough test planning, comprehensive requirements review, focus on business process testing, strong master data validation, rigorous regression testing, early environment preparation, user involvement in UAT."
    ),
    (
        "What are SAP testing challenges?",
        ["System complexity across many integrated modules",
         "Configuration customisation unique to each client",
         "Large data volumes needed for realistic testing",
         "Environment limitations and availability",
         "Timeline pressure during implementations"],
        "Challenges: system complexity (many integrated modules), configuration customisation (each client is unique), data volume (realistic data needed), environment limitations (not always available), change management (constant modifications), timeline pressure."
    ),
    (
        "Define SAP tester competencies.",
        ["SAP system knowledge: architecture, modules, transactions",
         "Business process understanding of how operations work",
         "Testing fundamentals: test design and defect management",
         "Problem-solving and analytical skills",
         "Domain knowledge of industry and functional area"],
        "Competencies: SAP system knowledge (architecture, modules, transactions), business process understanding (how business operates), testing fundamentals (test design, execution, defect management), problem-solving, communication, domain knowledge (industry/functional area)."
    ),
    (
        "What is SAP post-go-live support testing?",
        ["Testing during hypercare and stabilisation period",
         "Focuses on production defects found by users",
         "Urgent fixes requiring rapid verification",
         "Workaround validation for known issues",
         "Continues through the full hypercare period"],
        "Post-go-live testing focuses on: production defects found by users, urgent fixes needing verification, workaround validation, user questions about system behaviour. Support testing continues through the hypercare period."
    ),
    (
        "What is the strategic importance of SAP Tester in implementation?",
        ["Validates configured system meets business requirements",
         "Ensures data is accurate and complete",
         "Confirms processes work end-to-end",
         "Verifies system is production-ready",
         "Minimises go-live risks and post-go-live issues"],
        "The SAP Tester is critical by validating that: configured system meets business requirements, data is accurate and complete, processes work end-to-end, system is production-ready, go-live risks are minimised, post-go-live issues are prevented, and business value is realised from SAP investment."
    ),
]

# ── 100 SAP Test Manager Q&A ──────────────────────────────────────────────
SAP_TEST_MANAGER_QA = [
    (
        "What is Tosca?",
        ["Model-based test automation platform by Tricentis",
         "Enables non-technical users to create automated tests",
         "Visual modelling rather than coding",
         "Reduces maintenance effort for SAP testing",
         "Widely used for SAP functional and regression testing"],
        "Tosca (Test Optimization and Selection CAtech) is a model-based test automation platform used for automating functional and regression testing across enterprise applications including SAP. Tosca enables non-technical users to create automated tests through visual modelling rather than coding, reducing maintenance effort and improving test quality."
    ),
    (
        "Define Tosca's model-based approach.",
        ["Test automation using reusable components and modelling",
         "Test scenarios represented as executable models",
         "Built from reusable components (TOSCA Scans)",
         "Reduces redundancy and improves maintainability",
         "Allows rapid test creation and reuse"],
        "Tosca uses model-based testing where test scenarios are represented as executable models. Tests are built from reusable components (TOSCA Scans) mapped to application elements. This approach reduces redundancy, improves maintainability, and allows rapid test creation."
    ),
    (
        "What is SAP HANA?",
        ["High-Performance Analytic Appliance — SAP's in-memory database",
         "Stores data in RAM for extremely fast processing",
         "Enables real-time analytics",
         "Eliminates traditional batch processing delays",
         "Foundation of S/4HANA modern SAP"],
        "SAP HANA (High-Performance Analytic Appliance) is SAP's in-memory database technology that stores data in RAM for extremely fast processing. HANA enables real-time analytics, eliminating traditional batch processing delays. HANA is the foundation of S/4HANA (modern SAP)."
    ),
    (
        "Explain HANA's in-memory technology.",
        ["Data stored and processed in RAM, not disk",
         "Eliminates disk I/O bottlenecks",
         "Processing of massive datasets in seconds",
         "Enables real-time reporting without batch delays",
         "Fundamentally changes performance capabilities"],
        "HANA stores data in main memory (RAM) instead of disk storage. This eliminates disk I/O bottlenecks, enabling processing of massive datasets in seconds rather than minutes/hours. In-memory architecture fundamentally changes performance and reporting capabilities."
    ),
    (
        "What is the role of SAP Test Manager (with Tosca)?",
        ["Plans and manages automated testing strategy with Tosca",
         "Designs test strategy using Tosca capabilities",
         "Manages Tosca test execution and results",
         "Optimises test coverage and automation ROI",
         "Leads testing team using Tosca platform"],
        "SAP Test Manager using Tosca oversees test automation: designing test strategies using Tosca's capabilities, managing Tosca test execution, analysing results, optimising test coverage, and managing the testing team using Tosca platform."
    ),
    (
        "Define Tosca Test Suite.",
        ["Collection of related test cases grouped together",
         "Enables batch execution of related tests",
         "Results aggregated for comprehensive reporting",
         "Organised by process, module, or functionality",
         "Reused across test cycles"],
        "Tosca Test Suite groups related test cases together for organisation and execution. Suites can be executed together, enabling batch testing of related functionality. Suite results are aggregated for comprehensive reporting."
    ),
    (
        "What is Tosca Commander?",
        ["Central hub for managing all Tosca automation",
         "Primary interface for test execution",
         "Manages test assets and configurations",
         "Reviews results and configures integrations",
         "Provides centralised visibility into testing activities"],
        "Tosca Commander is the primary interface where testers and managers: execute tests, review results, manage test assets, configure integrations, and report on test metrics. Commander provides centralised visibility into testing activities."
    ),
    (
        "Explain Tosca Scan technology.",
        ["Intelligently identifies application UI elements",
         "Buttons, fields, links, tables all recognised",
         "Adapts to UI changes, reducing maintenance",
         "Smart object recognition improves reliability",
         "Creates mappings used across test cases"],
        "Tosca Scans intelligently identify application UI elements (buttons, fields, links) and create mappings. Scans are smart — they adapt to UI changes, reducing maintenance when the application changes. Smart object recognition improves test reliability."
    ),
    (
        "What is a Tosca Scan in SAP context?",
        ["Automated discovery of SAP screen elements",
         "Identifies SAP GUI fields, buttons, tables, dropdowns",
         "Creates object repositories for each element",
         "Quality scans are critical for test reliability",
         "Foundation for all SAP GUI automation"],
        "In SAP, Tosca Scans identify SAP GUI elements: input fields, buttons, tables, dropdown lists. Scans create object repositories mapping each element to Tosca components. Quality scans are critical for test reliability."
    ),
    (
        "Define Tosca object repository.",
        ["Central store of all recognised application elements",
         "Stores all SAP fields, buttons, and controls",
         "Enables reuse — many tests reference same elements",
         "Changes to element update all tests using it",
         "Foundation of maintainable automation"],
        "The object repository stores all recognised application elements (SAP fields, buttons, etc.) used in tests. Centralised repository enables reuse — many tests can reference same elements. Changes to element properties update all tests using that element."
    ),
    (
        "What is Tosca test execution?",
        ["Running automated tests and recording results",
         "Follows test steps and validates expected results",
         "Captures screenshots and logs",
         "Records pass/fail status for each step",
         "Can run unattended (scheduled) or attended"],
        "Test execution runs Tosca tests against SAP system: following test steps, validating expected results, capturing screenshots/logs, recording pass/fail status. Tosca can execute in unattended mode (scheduled) or attended mode (supervised)."
    ),
    (
        "Explain Tosca test result analysis.",
        ["Reviewing test execution outcomes",
         "Identifies which tests passed and which failed",
         "Investigates root causes of failures",
         "Distinguishes real bugs from test or environment issues",
         "Data-driven insights guide corrective action"],
        "Result analysis reviews: which tests passed/failed, what caused failures, performance metrics, coverage achieved. Analysis identifies: real bugs (defects in SAP), test issues (incorrect test design), and environment issues (system problems)."
    ),
    (
        "What is Tosca analytics dashboard?",
        ["Visualisation of test metrics and execution status",
         "Displays execution trends and defect identification rate",
         "Shows coverage progress and resource utilisation",
         "Enables data-driven testing management",
         "Provides quick problem identification"],
        "Tosca analytics dashboards display: test execution trends, defect identification rate, coverage progress, resource utilisation. Dashboards enable data-driven testing management and quick problem identification."
    ),
    (
        "Define Tosca integration capabilities.",
        ["Connects Tosca with other testing and development tools",
         "Defect tracking: JIRA, HP QC",
         "Requirements: ALM, Azure DevOps",
         "CI/CD: Jenkins, Azure DevOps pipelines",
         "Integration enables end-to-end process automation"],
        "Tosca integrates with: defect tracking (JIRA, HP QC), requirements (ALM, Azure DevOps), CI/CD (Jenkins, Azure DevOps), and reporting tools. Integration enables end-to-end testing process automation."
    ),
    (
        "What is HANA database testing?",
        ["Validating HANA-specific functionality and performance",
         "Data loads correctly into HANA",
         "In-memory processing works as expected",
         "Real-time analytics functions correctly",
         "Query performance meets expectations"],
        "HANA testing verifies: data loads correctly, in-memory processing works, real-time analytics functions, query performance meets expectations, HANA-specific features (columnar storage, compression) operate correctly."
    ),
    (
        "Explain HANA data volume testing.",
        ["Testing HANA with large datasets",
         "Millions of rows and terabytes of data",
         "Validates HANA handles scale effectively",
         "In-memory performance remains acceptable",
         "No data loss or corruption under volume"],
        "Volume testing pushes HANA with massive data: millions of rows, terabytes of data. Validates HANA handles scale, in-memory performance remains acceptable, and no data loss or corruption occurs."
    ),
    (
        "What is HANA performance tuning testing?",
        ["Validating impact of HANA optimisation",
         "Index optimisation improves query speed",
         "Partitioning distributes data efficiently",
         "Compression reduces memory usage",
         "Testing ensures tuning delivers expected benefits"],
        "Tuning testing verifies: index optimisation improves query speed, partitioning distributes data efficiently, compression reduces memory usage, calculated views perform well. Testing ensures tuning delivers expected benefits."
    ),
    (
        "Define S/4HANA.",
        ["SAP's next-generation ERP built on HANA database",
         "Simplified data model and real-time analytics",
         "Modern Fiori user interface",
         "Improved performance over classic SAP ECC",
         "SAP's strategic direction for ERP"],
        "S/4HANA is modern SAP ERP architecture built on HANA database. S/4HANA offers: simplified data model, real-time analytics, modern UI (Fiori), and improved performance. S/4HANA represents SAP's strategic direction."
    ),
    (
        "What is Tosca code-less approach?",
        ["Creating automated tests without programming",
         "Visual modelling and drag-and-drop interface",
         "Enables non-programmers to create tests",
         "Business analysts can create automation",
         "Democratises test automation across teams"],
        "Tosca enables non-programmers to create automated tests through visual modelling and drag-drop interface rather than coding. Code-less approach democratises test automation, enabling business analysts to create tests."
    ),
    (
        "Explain Tosca scripting capabilities.",
        ["Supports scripting for complex automation scenarios",
         "Custom scripting beyond standard GUI interactions",
         "Conditional logic and loop construction",
         "Complements code-less approach",
         "Enables advanced automation not possible via GUI"],
        "While Tosca emphasises code-less, it supports scripting for complex scenarios. Tosca supports: custom scripting, conditional logic, loop construction. Scripting enables advanced automation beyond standard GUI interactions."
    ),
    (
        "What is Tosca data management?",
        ["Handling test data within Tosca",
         "Test data creation and management",
         "Data masking for privacy compliance",
         "Data environment configuration",
         "Ensures tests run with realistic, appropriately masked data"],
        "Tosca includes data management features: test data creation, data masking (privacy), data environment configuration. Proper data management ensures tests run with realistic, appropriately masked data."
    ),
    (
        "Define Tosca test parameters.",
        ["Variables making tests flexible and reusable",
         "Same test runs with different data inputs",
         "Different customers, products, or quantities",
         "Increases test coverage without new test cases",
         "Supports iterative and data-driven testing"],
        "Tosca test parameters enable the same test to run with different data: different customers, products, quantities. Parameterisation increases test coverage without writing new tests. Parameters support iterative testing."
    ),
    (
        "What is Tosca distributed execution?",
        ["Running tests simultaneously across multiple machines",
         "Dramatically speeds test cycle time",
         "Requires appropriate licensing and infrastructure",
         "Enables parallel execution across environments",
         "Infrastructure management is required"],
        "Tosca supports distributed execution: tests run on multiple machines simultaneously, dramatically speeding test cycles. Distributed execution requires licence and infrastructure management."
    ),
    (
        "Explain Tosca licensing model.",
        ["Pricing and usage rights structure for Tosca",
         "Component licences for execution engine",
         "Pack licences for functional areas including SAP",
         "Execution licences for concurrent runs",
         "Licensing affects automation scope and scale"],
        "Tosca licensing includes: component licences (Tosca Core for execution engine), pack licences (functional areas like SAP), and execution licences. Licensing affects what can be automated and how many simultaneous executions."
    ),
    (
        "What is Tosca version management?",
        ["Tracking and controlling Tosca test asset changes",
         "Test assets are versioned and tracked",
         "Changes documented in version history",
         "Rollback to previous version possible",
         "Critical when multiple teams work on same suite"],
        "Version management ensures: test assets are versioned, changes are tracked, rollback is possible. Version control is critical when multiple teams work on the same test suite."
    ),
    (
        "Define HANA column store.",
        ["HANA's columnar storage architecture",
         "Data stored by column rather than row",
         "Improves compression significantly",
         "Accelerates analytical queries accessing specific columns",
         "HANA's performance foundation for analytics"],
        "HANA stores data by column (not row). Columnar storage: improves compression, accelerates analytical queries (which access specific columns), and enables faster data retrieval. Columnar architecture is HANA's performance foundation."
    ),
    (
        "What is HANA row store?",
        ["Row-based storage for transactional data in HANA",
         "Optimises OLTP transactions with small updates",
         "Complements column store for transactional workloads",
         "Some data uses row store, some column store",
         "Row store optimises real-time writes"],
        "While HANA's primary strength is column store, HANA also includes row store for transactional data. Row store optimises: OLTP transactions (small data updates), real-time writes. Some data uses row store, some column store."
    ),
    (
        "Explain HANA delta merge process.",
        ["Consolidating recent data changes with main store",
         "Merges recent changes into main data store",
         "Compresses data and optimises storage",
         "Improves query performance over time",
         "Automatic background process in HANA"],
        "HANA uses delta merge to consolidate recent changes into the main data store. Merge operations: compress data, optimise storage, improve query performance. Delta merge is an automatic background process."
    ),
    (
        "What is HANA compression?",
        ["Reducing storage requirement through encoding",
         "Dictionary encoding: common values replaced with codes",
         "Prefix encoding: repeated prefixes compressed",
         "High compression reduces memory requirement",
         "Reduces I/O and improves performance"],
        "HANA applies compression algorithms to reduce data size: dictionary encoding (common values replaced with codes), prefix encoding (repeated prefixes compressed). High compression reduces memory requirement and I/O."
    ),
    (
        "Define HANA calculated views.",
        ["Virtual tables computed on-the-fly from source data",
         "Compute values from source tables without storing results",
         "Enable complex calculations and aggregations",
         "Derived metrics without materialising data",
         "Foundation for HANA analytical modelling"],
        "Calculated views are virtual objects that compute values from source tables without storing results. Calculated views enable: complex calculations, aggregations, derived metrics without materialising data."
    ),
    (
        "What is the role of SAP Test Manager in Tosca implementation?",
        ["Leads Tosca adoption and automation strategy",
         "Tool selection, procurement, and licensing",
         "Team training on Tosca",
         "Test asset creation and maintenance",
         "Ensures Tosca delivers measurable ROI"],
        "SAP Test Manager leads: Tosca tool selection/procurement, team training, automation strategy development, test asset creation/maintenance, result analysis, and continuous improvement. Manager ensures Tosca delivers ROI."
    ),
    (
        "Explain Tosca ROI (Return on Investment).",
        ["Measuring value delivered by Tosca automation",
         "Cost savings from reduced manual testing",
         "Increased coverage and faster test cycles",
         "Compared against licensing, training, infrastructure costs",
         "Positive ROI typically achieved in 6-12 months"],
        "Tosca ROI calculation: cost savings from reduced manual testing + increased coverage + faster cycles vs. Tosca licensing + training + infrastructure. Positive ROI typically achieved in 6-12 months."
    ),
    (
        "What is Tosca test asset library?",
        ["Repository of reusable test components",
         "Contains TOSCA Scans, test cases, test modules",
         "Also test data and shared utilities",
         "Enables rapid test creation and reuse",
         "Reduces duplication and improves consistency"],
        "Asset library accumulates: TOSCA Scans (element mappings), test cases, test modules, test data, shared utilities. Robust library enables rapid test creation, reduces duplication, and improves consistency."
    ),
    (
        "Define Tosca modular testing.",
        ["Building tests from reusable functional modules",
         "Business processes decomposed into testable modules",
         "Modules tested independently and combined",
         "Faster development and easier maintenance",
         "Better organisation across large test suites"],
        "Tosca modular approach: business processes decomposed into testable modules, modules are tested independently and combined. Modularity enables: faster development, easier maintenance, better organisation."
    ),
    (
        "What is Tosca CI/CD integration?",
        ["Connecting Tosca to continuous integration/deployment pipelines",
         "Tests execute automatically on code commits",
         "Results feed deployment decision-making",
         "Enables continuous quality validation",
         "Supports DevOps and rapid delivery"],
        "Tosca integrates with CI/CD pipelines: tests execute automatically on code commits, results feedback to deployment decisions. CI/CD integration enables continuous quality validation."
    ),
    (
        "Explain Tosca API testing.",
        ["Automating interface testing in Tosca",
         "REST and SOAP service automation",
         "Request and response validation",
         "Performance measurement for APIs",
         "Extends beyond GUI to comprehensive service validation"],
        "Tosca includes API testing capabilities: REST/SOAP service automation, request/response validation, performance measurement. API testing extends beyond GUI to comprehensive service validation."
    ),
    (
        "What is HANA live data?",
        ["Real-time data without batching delays",
         "Updated in real-time, not batched",
         "Reports reflect current state instantly",
         "No refresh delays for business decisions",
         "Strategic competitive advantage for real-time decisions"],
        "HANA live means data is updated in real-time (not batched). Reporting on live data means: reports reflect current state, no refresh delays, business can make decisions on immediate data. Real-time is a strategic advantage."
    ),
    (
        "Define HANA analytics engine.",
        ["Processing engine optimised for analytical queries",
         "Handles complex queries across billions of rows",
         "Multi-dimensional analysis and aggregations",
         "In-memory architecture delivers millisecond performance",
         "Foundation for real-time business intelligence"],
        "HANA analytics engine optimises: complex queries across billions of rows, multi-dimensional analysis, aggregations. Engine's in-memory architecture executes analytical queries in milliseconds rather than minutes."
    ),
    (
        "What is HANA federation?",
        ["Connecting HANA with external data sources",
         "Query remote databases from HANA",
         "Join HANA data with remote data",
         "Analyse across multiple systems",
         "Eliminates data replication while enabling unified analysis"],
        "HANA federation enables querying: remote databases, cloud systems, Hadoop. Federated queries bring external data into HANA analysis without replicating to HANA."
    ),
    (
        "Explain Tosca continuous testing strategy.",
        ["Embedding automated testing throughout development",
         "Tests execute on every build",
         "Quick feedback on quality",
         "Early defect detection across all commits",
         "Shifts testing left, catching issues earlier"],
        "Continuous testing with Tosca means: tests execute on every build, quick feedback on quality, early defect detection. Continuous testing shifts testing left, catching issues earlier."
    ),
    (
        "What is Tosca risk-based testing?",
        ["Prioritising tests by business and technical risk",
         "Focus on critical business processes",
         "Complex and historically defect-prone areas",
         "Optimises limited testing resources",
         "Ensures highest-risk areas get most coverage"],
        "Risk-based testing focuses Tosca test suite on high-risk areas: critical business processes, complex functionality, areas with history of defects. Risk-based approach optimises limited testing resources."
    ),
    (
        "Define Tosca requirements traceability.",
        ["Linking tests to business and system requirements",
         "Each test traces to specific requirements",
         "Each requirement has associated tests",
         "Ensures requirements are tested",
         "Coverage is visible and auditable"],
        "Tosca enables traceability: each test traces to requirements, each requirement has associated tests. Traceability ensures: requirements are tested, tests are needed, coverage is evident."
    ),
    (
        "What is Tosca test case maintenance?",
        ["Updating tests when the application changes",
         "Intelligent object recognition reduces maintenance burden",
         "Centralised object repository means updates apply everywhere",
         "When UI changes, Tosca often adapts automatically",
         "Lower maintenance cost than traditional scripting"],
        "Tosca's intelligent object recognition reduces maintenance: when UI changes, Tosca often adapts automatically. When changes require test updates, Tosca's centralised object repository means updates apply everywhere."
    ),
    (
        "Explain Tosca reporting capabilities.",
        ["Generating insights from test results",
         "Test execution summaries and defect analysis",
         "Coverage metrics and trend reporting",
         "Custom reports for different audiences",
         "Enables data-driven decisions and stakeholder communication"],
        "Tosca reporting includes: test execution summaries, defect analysis, coverage metrics, trend reporting, custom reports. Rich reporting enables data-driven decisions and stakeholder communication."
    ),
    (
        "What is HANA partitioning?",
        ["Dividing large tables across storage locations",
         "Splits large tables into smaller parts",
         "Distributed across storage for parallel processing",
         "Improves query performance on very large tables",
         "Reduces memory footprint per partition"],
        "HANA partitioning: splits large tables into smaller parts, distributed across storage, enabling parallel processing. Partitioning improves query performance on very large tables."
    ),
    (
        "Define HANA replication.",
        ["Copying HANA data for backup and failover",
         "Maintains synchronous or asynchronous copies",
         "Enables failover to replica if primary fails",
         "Supports geo-distributed deployment",
         "Ensures business continuity"],
        "HANA replication: maintains synchronous or asynchronous copies of data, enables failover to replica if primary fails, supports geo-distributed deployment. Replication ensures business continuity."
    ),
    (
        "What is HANA smart data access?",
        ["Querying external data as if it were in HANA",
         "Query remote databases directly from HANA",
         "Join HANA data with remote data sources",
         "Analyse across systems without replication",
         "Eliminates need to replicate external data into HANA"],
        "Smart data access enables: querying remote databases directly from HANA, joining HANA data with remote data, analysing across systems. Eliminates data replication while enabling unified analysis."
    ),
    (
        "Explain Tosca test orchestration.",
        ["Coordinating complex multi-step test scenarios",
         "Combines multiple test cases into workflows",
         "Executes scenarios in sequence",
         "Passes data between test steps",
         "Enables end-to-end process testing"],
        "Test orchestration: combines multiple test cases into workflows, executes scenarios in sequence, passes data between steps. Orchestration enables end-to-end process testing."
    ),
    (
        "What is Tosca parallel execution?",
        ["Running multiple tests simultaneously",
         "Different test cases on different machines",
         "Different scenarios in parallel",
         "Dramatically reduces overall test cycle time",
         "Requires distributed infrastructure"],
        "Parallel execution runs: different test cases on different machines, different scenarios in parallel. Parallel execution dramatically reduces overall test cycle time."
    ),
    (
        "Define SAP Test Manager responsibilities with Tosca/HANA.",
        ["Oversees automated testing in SAP HANA environment",
         "Tosca automation strategy and governance",
         "HANA test environment management",
         "Test suite design, maintenance, and analysis",
         "Stakeholder communication and continuous improvement"],
        "Responsibilities include: Tosca automation strategy, HANA test environment management, test suite design and maintenance, result analysis, team leadership, continuous improvement, and stakeholder communication."
    ),
    (
        "What is HANA data load testing?",
        ["Validating ETL processes into HANA",
         "Executes Extract, Transform, Load processes",
         "Validates data accuracy after loading",
         "Verifies completeness — no data lost",
         "Checks ETL performance is acceptable"],
        "Data load testing: executes ETL (Extract, Transform, Load) processes, validates data accuracy, verifies completeness, checks performance. Load testing ensures the data pipeline works correctly."
    ),
    (
        "Explain Tosca keyword-driven testing.",
        ["Using business-friendly language for automation",
         "Tests written using business keywords",
         "Enables non-technical people to understand tests",
         "Keywords map to technical implementation",
         "Bridges business and technical teams"],
        "Keyword-driven approach: tests written using business keywords rather than technical code, enabling non-technical people to understand/modify tests. Keywords map to technical implementation."
    ),
    (
        "What is Tosca behaviour-driven development (BDD)?",
        ["Using behaviour specifications for test automation",
         "Tests written in Gherkin language",
         "Stakeholders can review test specifications",
         "Developers implement specified behaviour",
         "Tests automate the specifications"],
        "BDD in Tosca: tests written in Gherkin language describing behaviour, stakeholders can review test specifications, developers implement behaviour, tests automate the specifications."
    ),
    (
        "Define HANA query optimisation.",
        ["Tuning queries for best performance in HANA",
         "Selecting right index for query",
         "Optimal join strategy selection",
         "Appropriate aggregation method",
         "SAP Test Manager validates no performance regression"],
        "Query optimisation: selecting right index, optimal join strategy, appropriate aggregation method. HANA test manager validates: queries execute efficiently, indexes are properly designed, no performance regression."
    ),
    (
        "What is Tosca test case template?",
        ["Predefined structure ensuring consistent test cases",
         "Standardises steps, expected results, data requirements",
         "Ensures consistency across the team",
         "Eases training for new team members",
         "Enables faster test creation"],
        "Templates standardise test case format: steps, expected results, data requirements. Templates ensure consistency, ease training, and enable faster test creation."
    ),
    (
        "Explain SAP Test Manager's role in Tosca training.",
        ["Educates team on Tosca capabilities and best practices",
         "Tosca basics: navigation and test creation",
         "SAP-specific testing with Tosca",
         "Best practices and troubleshooting",
         "Well-trained team leverages Tosca effectively"],
        "Training includes: Tosca basics (navigation, test creation), SAP-specific testing with Tosca, best practices, troubleshooting. A well-trained team leverages Tosca effectively."
    ),
    (
        "What is Tosca monitoring during execution?",
        ["Watching tests run in real-time and detecting issues",
         "Identifies failures as they occur",
         "Enables intervention if needed",
         "Collects performance data during execution",
         "Enables rapid issue response"],
        "Monitoring: watches test execution in real-time, identifies failures as they occur, enables intervention if needed, collects performance data. Real-time monitoring enables rapid issue response."
    ),
    (
        "Define HANA security testing.",
        ["Validating access controls and data protection in HANA",
         "Authentication works correctly",
         "Authorisation restricts data access appropriately",
         "Encryption protects sensitive data",
         "Audit logs record all relevant activity"],
        "Security testing: verifies authentication works, authorisation restricts data access, encryption protects data, audit logs record activity. Security testing ensures HANA security posture is strong."
    ),
    (
        "What is Tosca SAP GUI testing?",
        ["Automating traditional SAP GUI transactions",
         "Recognises SAP GUI elements",
         "Automates transaction entry and navigation",
         "Validates results and postings",
         "Covers traditional SAP functionality not yet in Fiori"],
        "Tosca SAP GUI testing: recognises SAP GUI elements, automates transaction entry, validates results. SAP GUI testing covers traditional SAP functionality not yet in Fiori."
    ),
    (
        "Explain Tosca Fiori testing.",
        ["Automating modern SAP Fiori UI interactions",
         "Recognises Fiori app elements and controls",
         "Handles responsive design testing",
         "Automates touch gestures for mobile",
         "Addresses next-generation SAP UI"],
        "Fiori testing in Tosca: recognises Fiori app elements, handles responsive design, automates touch gestures (if mobile), validates modern UI. Fiori testing addresses next-generation SAP UI."
    ),
    (
        "What is SAP Test Manager's role in test environment management?",
        ["Ensuring appropriate test infrastructure",
         "Maintains multiple test environments: Dev, QA, UAT",
         "Configuration consistency across environments",
         "Data refresh capability",
         "Environment availability when teams need them"],
        "Environment management: maintains multiple test environments (Dev, QA, UAT), configuration consistency across environments, data refresh capability, environment availability. Proper environments enable effective testing."
    ),
    (
        "Define Tosca business process analytics.",
        ["Understanding test coverage and defect patterns by process",
         "Identifies which business processes have best test coverage",
         "Which areas have most defects",
         "Which tests find most issues",
         "Data-driven insights optimise testing focus"],
        "Analytics: identifies: which business processes have best test coverage, which areas have most defects, which tests find most issues, where to focus testing. Data-driven insights optimise testing."
    ),
    (
        "What is HANA advanced analytics?",
        ["Analytics capabilities beyond traditional BI",
         "Predictive analytics: forecasting future trends",
         "Clustering analysis: grouping similar data",
         "Anomaly detection: identifying unusual patterns",
         "Enables sophisticated analytical models"],
        "Advanced analytics: predictive analytics (forecasting future trends), clustering analysis (grouping similar customers), anomaly detection (identifying unusual patterns). HANA enables sophisticated analytical models."
    ),
    (
        "Explain Tosca test automation roadmap.",
        ["Strategic plan for progressive automation",
         "Phases automation over time",
         "Prioritises high-value areas for automation",
         "Plans resource and tool requirements",
         "Ensures sustainable automation aligned with business needs"],
        "Roadmap: phases automation over time, prioritises high-value areas, plans resource/tool needs, defines milestones. Roadmap ensures sustainable automation growth aligned with business needs."
    ),
    (
        "What is Tosca defect auto-detection?",
        ["Automatically identifying failures in test results",
         "Compares expected vs actual results",
         "Detects mismatches automatically",
         "Reports defects without manual verification",
         "Reduces manual effort and improves consistency"],
        "Auto-detection: compares expected vs. actual results, detects mismatches automatically, reports defects. Auto-detection reduces manual verification effort and improves consistency."
    ),
    (
        "Define HANA MDC (Multi-Database Containers).",
        ["Running multiple isolated HANA databases in one installation",
         "Multiple HANA systems isolated as tenants",
         "Resources shared but isolated per tenant",
         "Enables consolidation while maintaining isolation",
         "Useful for multi-client hosting environments"],
        "MDC allows: multiple HANA systems in single installation, isolated tenants, resource sharing. MDC enables consolidation while maintaining isolation — useful for hosting providers."
    ),
    (
        "What is SAP Test Manager's role in vendor management?",
        ["Overseeing Tosca vendor relationship",
         "Maintains licence compliance",
         "Manages support tickets with Tricentis",
         "Negotiates licence renewals",
         "Stays updated on new Tosca capabilities"],
        "Vendor management: maintains licence compliance, manages support tickets, negotiates renewals, stays updated on new capabilities. Strong vendor relationship ensures effective tool usage."
    ),
    (
        "Explain Tosca sanity testing.",
        ["Quick automated verification after deployments",
         "Verifies basic functionality after changes",
         "Identifies major issues immediately",
         "Runs before detailed testing begins",
         "Fail fast on critical issues"],
        "Sanity testing with Tosca: quick automated tests verify basic functionality, identifies major issues before detailed testing. Sanity tests run first, fail fast on critical issues."
    ),
    (
        "What is Tosca smoke testing in SAP?",
        ["Brief automated verification after SAP updates",
         "Confirms SAP basics still work after configuration changes",
         "Identifies critical problems quickly",
         "Runs first in every test cycle",
         "Quick early detection of breaking changes"],
        "Smoke tests: automated verification that SAP basics still work after configuration changes/patches, quick early detection of problems. Smoke tests run first in the test cycle."
    ),
    (
        "Define HANA streaming analytics.",
        ["Real-time analysis of event data streams",
         "Ingests continuous event streams",
         "Analyses in real-time",
         "Identifies patterns and anomalies immediately",
         "Enables real-time alerting and response"],
        "Streaming analytics: ingests continuous event streams, analyses in real-time, identifies patterns/anomalies immediately. Streaming enables real-time alerting and response."
    ),
    (
        "What is Tosca test management with ALM integration?",
        ["Connecting Tosca with requirements and test management tools",
         "Tests link to requirements in ALM",
         "Requirements traceability maintained",
         "Test status updates requirements view",
         "Enables comprehensive lifecycle management"],
        "ALM integration: tests link to requirements, requirements traceability maintained, test status updates requirements view. Integration enables comprehensive lifecycle management."
    ),
    (
        "Explain SAP Test Manager's role in defect management.",
        ["Overseeing defect identification and resolution",
         "Tosca identifies and logs defects",
         "Test Manager categorises defects: real bug vs test issue",
         "Prioritises and facilitates resolution",
         "Ensures systematic defect handling"],
        "Defect management: Tosca identifies defects, test manager categorises (real bug vs. test issue), prioritises, facilitates resolution, tracks closure. Manager ensures systematic defect handling."
    ),
    (
        "What is Tosca exploratory testing support?",
        ["Enabling flexible manual testing alongside automation",
         "Testers manually explore while automation covers core flows",
         "Combination of automated and manual optimises coverage",
         "Exploratory testing finds what automation misses",
         "Hybrid approach is most comprehensive"],
        "While Tosca focuses automation, it supports exploratory testing: testers can manually explore while automating core flows. Combination of automated + manual optimises coverage."
    ),
    (
        "Define HANA data modelling.",
        ["Designing optimal data structures for HANA",
         "Designing column vs row store usage",
         "Optimising for analytical vs transactional access",
         "Defining aggregations and partitioning",
         "Good modelling ensures performance"],
        "HANA data modelling: designing column vs. row store usage, optimising for analytical vs. transactional access, defining aggregations, partitioning large tables. Good modelling ensures performance."
    ),
    (
        "What is Tosca test case versioning?",
        ["Managing test case changes over time",
         "Tracks all test case modifications",
         "Enables rollback to prior versions",
         "Documents change history",
         "Critical in team testing environments"],
        "Versioning: tracks test case modifications, enables rollback to prior versions, documents change history. Version control is critical in team environments."
    ),
    (
        "Explain SAP Test Manager's role in quality gates.",
        ["Enforcing quality standards for SAP releases",
         "Defines criteria for release approval",
         "Test coverage %, defects closed, performance SLAs",
         "Enforces criteria before go-live approval",
         "Gates prevent unready releases reaching production"],
        "Quality gates: define criteria for release approval (test coverage %, defects closed, performance SLAs met), enforce criteria before go-live. Gates prevent unready releases."
    ),
    (
        "What is Tosca scalability in large SAP environments?",
        ["Handling automation in complex enterprise organisations",
         "Supports hundreds of test cases",
         "Multiple teams and distributed execution",
         "Handles massive datasets",
         "Enables enterprise-wide automation"],
        "Tosca scales: supports hundreds of test cases, multiple teams, distributed execution, massive datasets. Scalability enables enterprise-wide automation."
    ),
    (
        "Define HANA predictive analytics.",
        ["Using historical data for forecasting and prediction",
         "Builds models from historical data",
         "Predicts future trends and outcomes",
         "Enables proactive business decisions",
         "Provides competitive advantage through foresight"],
        "Predictive analytics: builds models from historical data, predicts future trends/outcomes, enables proactive decisions. Predictive capabilities provide competitive advantage."
    ),
    (
        "What is Tosca compliance testing?",
        ["Validating regulatory requirement adherence",
         "Verifies SOX, GDPR, HIPAA requirements are met",
         "Audit trails function correctly",
         "Access controls work as required",
         "Provides assurance for regulated industries"],
        "Compliance testing: verifies system meets regulatory requirements (SOX, GDPR, HIPAA), audit trails function correctly, access controls work. Compliance testing provides assurance for regulated industries."
    ),
    (
        "Explain SAP Test Manager's role in continuous improvement.",
        ["Driving process optimisation in testing",
         "Analyses test metrics for bottlenecks",
         "Implements process enhancements",
         "Measures impact of improvements",
         "Increases testing maturity over time"],
        "Continuous improvement: analyses test metrics, identifies bottlenecks, implements process enhancements, measures impact. Focus on improvement increases testing maturity."
    ),
    (
        "What is Tosca integration testing across systems?",
        ["Testing data flow between SAP and external systems",
         "Validates data transfers correctly",
         "Formats match specifications",
         "No data loss occurs during transfer",
         "Timing and frequency are appropriate"],
        "Integration testing: Tosca validates: data transfers correctly, formats match specifications, no loss occurs, timing is appropriate. Integration testing catches interface issues early."
    ),
    (
        "Define HANA data governance.",
        ["Controlling data quality and usage in HANA",
         "Establishes data quality standards",
         "Defines ownership and responsibility",
         "Sets usage policies and compliance",
         "Master data management discipline"],
        "Data governance: establishes data quality standards, ownership responsibility, usage policies, master data management. Governance ensures trustworthy data."
    ),
    (
        "What is Tosca test failure analysis?",
        ["Root cause analysis of automated test failures",
         "Investigates why specific tests failed",
         "Distinguishes real defects from test issues",
         "Identifies patterns in failures",
         "Guides corrective action and remediation"],
        "Failure analysis: investigates why tests failed, distinguishes real defects from test issues, identifies patterns. Analysis guides corrective action."
    ),
    (
        "Explain SAP Test Manager's role in stakeholder reporting.",
        ["Communicating testing status and insights to stakeholders",
         "Executives receive quality status overview",
         "Project managers receive progress updates",
         "Business receives functionality validation confirmation",
         "Tailored reporting for each audience"],
        "Reporting: executives get quality status, project managers get progress updates, business gets functionality validation confirmation. Tailored reporting keeps stakeholders informed."
    ),
    (
        "What is Tosca accessibility testing?",
        ["Automating compliance with accessibility standards",
         "Validates UI meets accessibility requirements (WCAG)",
         "Keyboard navigation works correctly",
         "Screen reader compatibility confirmed",
         "Ensures inclusive design compliance"],
        "Accessibility testing: validates UI meets accessibility requirements (WCAG), keyboard navigation works, screen reader compatibility. Accessibility testing ensures inclusive design."
    ),
    (
        "Define HANA business continuity.",
        ["Ensuring HANA availability and recovery capability",
         "Replication maintains data copies",
         "Failover enables rapid recovery",
         "Backup and restore capability verified",
         "Disaster recovery planning and testing"],
        "Business continuity: replication maintains data copies, failover enables rapid recovery, backup/restore capability, disaster recovery planning. Continuity ensures business resilience."
    ),
    (
        "What is Tosca mobile app testing?",
        ["Automating tests for SAP mobile applications",
         "Handles touch gestures and mobile interactions",
         "Responsive design tested across screen sizes",
         "Mobile-specific scenarios covered",
         "Addresses SAP's expanding mobile application portfolio"],
        "Mobile testing: Tosca handles touch gestures, responsive design, mobile-specific scenarios. Mobile testing addresses SAP's expanding mobile applications."
    ),
    (
        "Explain SAP Test Manager's role in innovation.",
        ["Adopting new testing approaches and technologies",
         "Evaluates new testing tools and methods",
         "Explores emerging methodologies",
         "Experiments with advanced techniques",
         "Drives organisational testing evolution"],
        "Innovation: evaluates new testing tools, explores emerging methodologies, experiments with advanced techniques, drives organisational evolution. Innovation keeps testing competitive."
    ),
    (
        "What is Tosca industry-specific testing?",
        ["Adapting Tosca for specific domain requirements",
         "Retail: inventory and customer testing",
         "Financial: regulatory and compliance testing",
         "Manufacturing: production and quality testing",
         "Tosca adapts to industry-specific SAP configurations"],
        "Industry-specific testing: SAP implementations for retail require inventory testing, financial companies need regulatory testing, manufacturers need production testing. Tosca adapts to industry needs."
    ),
    (
        "Define HANA machine learning integration.",
        ["Applying ML models within HANA",
         "Trains models using HANA data and processing",
         "Predicts using trained models",
         "Operationalises algorithms in HANA",
         "HANA's performance enables sophisticated ML applications"],
        "ML integration: trains models in HANA, predicts using models, operationalises algorithms. HANA's performance enables sophisticated ML applications."
    ),
    (
        "What is SAP Test Manager's role in Agile/DevOps environments?",
        ["Enables testing in rapid delivery cycles",
         "Rapid test automation for each sprint",
         "Continuous testing integration in pipelines",
         "Quick feedback loops for developers",
         "Adapts testing to agile cadence and sprint goals"],
        "Agile/DevOps: rapid test automation, continuous testing integration, quick feedback loops, sprint-based test planning. Manager adapts testing to agile cadence."
    ),
    (
        "Explain Tosca hybrid testing approach.",
        ["Combining automated and manual testing",
         "Automates repetitive and regression tests",
         "Reserves manual for exploratory and complex scenarios",
         "Combines results for comprehensive view",
         "Balances efficiency with depth of coverage"],
        "Hybrid approach: automates repetitive tests, reserves manual for exploratory/complex scenarios, combines results for comprehensive view. Hybrid balances efficiency and coverage."
    ),
    (
        "What is HANA edge case handling?",
        ["Testing boundary conditions and unusual scenarios",
         "Zero amounts handled correctly",
         "Negative numbers processed properly",
         "Maximum limits respected",
         "Special characters supported throughout"],
        "Edge case testing: validates: zero amounts handled, negative numbers processed, maximum limits respected, special characters supported. Edge case testing prevents production surprises."
    ),
    (
        "Define SAP Test Manager's strategic impact.",
        ["Long-term organisational testing transformation",
         "Builds lasting testing capability and maturity",
         "Drives quality culture across organisation",
         "Enables continuous delivery and deployment",
         "Supports digital transformation with quality"],
        "Strategic impact: builds testing capability, drives quality culture, enables continuous delivery, supports digital transformation. Strategic focus creates lasting value."
    ),
    (
        "What is Tosca test documentation practices?",
        ["Creating comprehensive, maintainable test documentation",
         "Describes test purpose and scope",
         "Step-by-step execution instructions",
         "Expected results clearly documented",
         "Maintenance notes for future team members"],
        "Documentation: describes test purpose, step-by-step execution, expected results, maintenance notes. Documentation enables knowledge transfer and supports audits."
    ),
    (
        "Explain HANA migration testing.",
        ["Validating transition from legacy system to HANA",
         "Data migrates accurately and completely",
         "New architecture works correctly",
         "Performance meets or exceeds expectations",
         "Rollback is possible if migration fails"],
        "Migration testing: verifies data migrates accurately, new architecture works correctly, performance meets expectations, rollback possible if needed. Migration testing ensures successful transition."
    ),
    (
        "What is Tosca test measurement and analytics?",
        ["Quantifying testing effectiveness with metrics",
         "Test coverage percentage tracked",
         "Defect detection rate measured",
         "Test execution time monitored",
         "Automation ROI calculated and reported"],
        "Analytics: measures: test coverage %, defect detection rate, test execution time, automation ROI. Metrics enable data-driven testing optimisation."
    ),
    (
        "Define SAP Test Manager's team development role.",
        ["Building high-performing testing teams",
         "Recruiting skilled SAP testers",
         "Technical training in Tosca, SAP, HANA",
         "Soft skills development and coaching",
         "Career pathing and development planning"],
        "Team development: recruiting skilled testers, technical training (Tosca, SAP, HANA), soft skills development, career pathing. Developed teams deliver superior results."
    ),
    (
        "What is Tosca optimisation for large test suites?",
        ["Improving efficiency with hundreds of automated tests",
         "Organising tests logically by process and risk",
         "Parallel execution to reduce cycle time",
         "Smart test selection: not running all tests every time",
         "Result analysis focus on high-value tests"],
        "Optimisation: organising tests logically, parallel execution, smart test selection (not running all every time), result analysis focus. Optimisation maintains suite manageability."
    ),
    (
        "What is the strategic importance of SAP Test Manager in Tosca/HANA environment?",
        ["Enables efficient Tosca automation deployment",
         "Manages quality in complex HANA environment",
         "Delivers measurable testing ROI",
         "Ensures successful SAP modernisation",
         "Supports business value realisation through quality"],
        "The SAP Test Manager is critical by: enabling efficient Tosca automation deployment, managing testing in HANA environment, delivering quality assurance, optimising testing ROI, ensuring successful SAP modernisation, supporting business value realisation through quality software."
    ),
]


def _seed_role(role: str, qa_list: list, label: str) -> None:
    print(f"\nSeeding {len(qa_list)} Q&A for role: {role}...")
    for idx, (question, points, illustrated) in enumerate(qa_list, start=1):
        excel_service.save_qa({
            "id": str(uuid.uuid4()),
            "role": role,
            "question": question,
            "points": points,
            "illustrated": illustrated,
            "source": "seed",
            "session_id": "seed-initial",
            "timestamp": datetime.now(timezone.utc).isoformat(),
        })
        print(f"  [{idx:03d}/{len(qa_list)}] {question[:70]}...")
    print(f"  [OK] {label} seeded.")


def seed():
    print("=" * 60)
    print("Interview Assistant - Excel Database Seeder")
    print("=" * 60)

    # Ensure all role sheets exist first
    from services import excel_service as es
    es._get_or_create_workbook()

    _seed_role("Test Manager",      TEST_MANAGER_QA,      "Test Manager (100 Q&A)")
    _seed_role("Scrum Master",      SCRUM_MASTER_QA,      "Scrum Master (100 Q&A)")
    _seed_role("Project Manager",   PROJECT_MANAGER_QA,   "Project Manager (100 Q&A)")
    _seed_role("Defect Manager",    DEFECT_MANAGER_QA,    "Defect Manager (100 Q&A)")
    _seed_role("SAP Tester",        SAP_TESTER_QA,        "SAP Tester (100 Q&A)")
    _seed_role("SAP Test Manager",  SAP_TEST_MANAGER_QA,  "SAP Test Manager (100 Q&A)")

    print("\n" + "=" * 60)
    print("Done! Excel file: interview-assistant/data/interview_qa_database.xlsx")
    print("Seeded roles:")
    print("  [OK] Test Manager      - 100 Q&A")
    print("  [OK] Scrum Master      - 100 Q&A")
    print("  [OK] Project Manager   - 100 Q&A")
    print("  [OK] Defect Manager    - 100 Q&A")
    print("  [OK] SAP Tester        - 100 Q&A")
    print("  [OK] SAP Test Manager  - 100 Q&A")
    print("\nTo start the server: python server.py")
    print("=" * 60)


if __name__ == "__main__":
    seed()
