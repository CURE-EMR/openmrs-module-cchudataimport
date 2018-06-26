package org.openmrs.module.cchudataimport.api.csv.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openmrs.Concept;
import org.openmrs.Encounter;
import org.openmrs.Obs;
import org.openmrs.Patient;
import org.openmrs.PatientIdentifier;
import org.openmrs.PatientIdentifierType;
import org.openmrs.PersonAttribute;
import org.openmrs.PersonAttributeType;
import org.openmrs.Visit;
import org.openmrs.api.APIException;
import org.openmrs.api.ConceptService;
import org.openmrs.api.PatientService;
import org.openmrs.api.context.Context;

public class SimpleExcelReader {
	
	/** Logger for this class and subclasses */
	protected final Log log = LogFactory.getLog(getClass());
	
	private String duplicateIds = "160892";
	
	public SimpleExcelReader() {
		super();
		// TODO Auto-generated constructor stub
	}
	
	public void prepareObsImport() throws APIException, IOException {
		ConceptService cs = Context.getConceptService();
		String mapping = Context.getAdministrationService().getGlobalProperty("cchudataimport.fileNameConceptMap");
		String[] values = mapping.split("\\|");
		for (String s : values) {
			if (s.indexOf(":") > 0) {
				String fileName = s.substring(0, s.indexOf(":"));
				String columnConcept = s.substring(s.indexOf(":") + 1, s.length());
				doImport(fileName, cs.getConcept(columnConcept));
			}
		}
		
	}
	
	public void preparePersonAttributesImport() throws APIException, IOException {
		ConceptService cs = Context.getConceptService();
		String mapping = Context.getAdministrationService().getGlobalProperty("cchudataimport.fileNamePersonAttributeMap");
		log.error("===================Ndi muri prepare kandi global property nayibonye. Ni: " + mapping);
		String[] values = mapping.split("\\|");
		for (String s : values) {
			if (s.indexOf(":") > 0) {
				String fileName = s.substring(0, s.indexOf(":"));
				PersonAttributeType personAttributeType = Context.getPersonService().getPersonAttributeTypeByName(s.substring(s.indexOf(":") + 1, s.length()));
				if (personAttributeType != null) {
					addPersonAttributes(fileName, personAttributeType);
				}
			}
		}
		
	}
	
	public void doImport(String fileName, Concept columnConcept) throws IOException {
		String excelFilePath = "/opt/openmrs/modules/" + fileName + ".xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();
		
		while (iterator.hasNext()) {
			try {
				Row nextRow = iterator.next();
				String oldEncounterUUID = nextRow.getCell(0).getStringCellValue();
				Cell c = nextRow.getCell(1);
				c.setCellType(Cell.CELL_TYPE_STRING);
				String obsValue = c.getStringCellValue();
				if (obsValue != null && obsValue.length() > 0 && !obsValue.equalsIgnoreCase("NULL") && !obsValue.equalsIgnoreCase("0000-00-00")) {
					createObs(oldEncounterUUID, obsValue, columnConcept);
				}
			}
			catch (Exception e) {
				// TODO: handle exception
			}
			
		}
		inputStream.close();
	}
	
	public Obs createObs(String oldEncounterUUID, String obsValue, Concept columnConcept) {
		ConceptService cs = Context.getConceptService();
		Concept c = cs.getConcept("Old Encounter UUID");
		List<Obs> obs = Context.getObsService().getObservationsByPersonAndConcept(null, c);
		Encounter e = null;
		for (Obs o : obs) {
			if (o.getValueText().equalsIgnoreCase(oldEncounterUUID)) {
				e = o.getEncounter();
			}
		}
		Obs o = new Obs();
		o.setConcept(columnConcept);
		if (columnConcept.getDatatype().isText()) {
			o.setValueText(obsValue);
		} else if (columnConcept.getDatatype().isCoded()) {
			Concept valueCoded = cs.getConcept(obsValue);
			if (valueCoded != null) {
				o.setValueCoded(valueCoded);
			}
			
		} else if (columnConcept.getDatatype().isDate()) {
			Date d = parseDate(obsValue);
			if (d != null) {
				o.setValueDate(d);
			}
		} else if (columnConcept.getDatatype().isNumeric()) {
			double value = parseToDouble(obsValue);
			if (value != -1) {
				o.setValueNumeric(value);
			}
			
		}
		int obsGroup = Integer.parseInt(Context.getAdministrationService().getGlobalProperty("cchudataimport.obsGroupConcept"));
		for (Obs obs2 : e.getObs()) {
			if (obs2.getConcept().getConceptId() == obsGroup) {
				o.setObsGroup(obs2);
			} else {
				Obs group = new Obs(e.getPatient(), cs.getConcept(obsGroup), e.getEncounterDatetime(), e.getLocation());
				e.addObs(group);
				o.setObsGroup(group);
			}
			
		}
		o.setObsDatetime(e.getEncounterDatetime());
		o.setLocation(e.getLocation());
		e.addObs(o);
		Context.getEncounterService().saveEncounter(e);
		return o;
	}
	
	public void addDateofBirth() throws IOException {
		String excelFilePath = "/opt/openmrs/modules/dof.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();
		
		PatientIdentifierType it = Context.getPatientService().getPatientIdentifierTypeByName("Old Identification Number");
		List<PatientIdentifierType> oldIds = new ArrayList<PatientIdentifierType>();
		oldIds.add(it);
		
		List<Patient> patients = Context.getPatientService().getAllPatients();//Do we really have to load all patients!!
		
		while (iterator.hasNext()) {
			try {
				Row nextRow = iterator.next();
				String patientUUID = nextRow.getCell(0).getStringCellValue();
				String dateOfBirth = nextRow.getCell(1).getStringCellValue();
				if (dateOfBirth != null && dateOfBirth.length() > 0 && !dateOfBirth.equalsIgnoreCase("NULL")) {
					Date dob = parseDate(dateOfBirth);
					if (dob != null) {
						for (Patient patient : patients) {
							if (patient.getPatientIdentifier(it).getIdentifier().equalsIgnoreCase(patientUUID)) {
								patient.setBirthdate(dob);
								Context.getPatientService().savePatient(patient);
							}
						}
					}
				}
			}
			catch (Exception e) {
				// TODO: handle exception
			}
			
		}
		inputStream.close();
	}
	
	public void addRegistrationDiagnosis() throws IOException {
		String excelFilePath = "/opt/openmrs/modules/registrationDiagnosis.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();
		
		PatientService ps = Context.getPatientService();
		PatientIdentifierType it = ps.getPatientIdentifierTypeByName("Old Identification Number");
		PersonAttributeType pt = Context.getPersonService().getPersonAttributeTypeByUuid("537f22bb-8b4e-4d51-9f54-d3b315a1a2d2");
		List<PatientIdentifierType> oldIds = new ArrayList<PatientIdentifierType>();
		oldIds.add(it);
		
		List<PatientIdentifier> identifiers = Context.getPatientService().getPatientIdentifiers(null, oldIds, null, null, null); //Do we really have to load all patients!!
		
		while (iterator.hasNext()) {
			try {
				Row nextRow = iterator.next();
				String patientUUID = nextRow.getCell(0).getStringCellValue();
				String registrationDiagnosis = nextRow.getCell(1).getStringCellValue();
				if (registrationDiagnosis != null && registrationDiagnosis.length() > 0 && !registrationDiagnosis.equalsIgnoreCase("NULL")) {
					for (PatientIdentifier patientIdentifier : identifiers) {
						if (patientIdentifier.getIdentifier().equalsIgnoreCase(patientUUID)) {
							Patient p = patientIdentifier.getPatient();
							Concept c = Context.getConceptService().getConceptByName(registrationDiagnosis);
							if (c != null) {
								PersonAttribute pa = new PersonAttribute(pt, c.getConceptId().toString());
								p.getPerson().addAttribute(pa);
								Context.getPersonService().savePerson(p);
							}
						}
					}
				}
			}
			catch (Exception e) {
				// TODO: handle exception
			}
			
		}
		inputStream.close();
	}
	
	public void addPersonAttributes(String fileName, PersonAttributeType personAttributeType) throws IOException {
		String excelFilePath = "/opt/openmrs/modules/" + fileName + ".xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();
		
		PatientService ps = Context.getPatientService();
		PatientIdentifierType it = ps.getPatientIdentifierTypeByName("Identification Number");
		List<PatientIdentifierType> oldIds = new ArrayList<PatientIdentifierType>();
		oldIds.add(it);
		
		List<PatientIdentifier> identifiers = Context.getPatientService().getPatientIdentifiers(null, oldIds, null, null, null); //Do we really have to load all patients!!
		log.error("================================Ndi muri addPersonAttributes kandi indentifiers nazibonye. Nabonye " + identifiers.size());
		while (iterator.hasNext()) {
			try {
				Row nextRow = iterator.next();
				String patientUUID = null;
				
				Cell c0 = nextRow.getCell(0);
				if (c0 != null) {
					c0.setCellType(Cell.CELL_TYPE_STRING);
					patientUUID = c0.getStringCellValue();
				}
				
				Cell c1 = nextRow.getCell(1);
				String personAttributeValue = null;
				
				if (c1 != null) {
					c1.setCellType(Cell.CELL_TYPE_STRING);
					personAttributeValue = c1.getStringCellValue();
				}
				
				if (patientUUID != null && patientUUID.length() > 0 && personAttributeValue != null && personAttributeValue.length() > 0 && !personAttributeValue.equalsIgnoreCase("NULL")) {
					for (PatientIdentifier patientIdentifier : identifiers) {
						if (patientIdentifier.getIdentifier().equalsIgnoreCase(patientUUID)) {
							Patient p = patientIdentifier.getPatient();
							PersonAttribute pa = new PersonAttribute(personAttributeType, personAttributeValue);
							p.getPerson().addAttribute(pa);
							log.error("======================================= " + patientUUID + "=============" + personAttributeType.getName());
							Context.getPersonService().savePerson(p);
						}
					}
				}
			}
			catch (Exception e) {
				log.error("Habaye ikibazo=======================" + e);
			}
			log.error("Ndacyasomaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa=======================");
		}
		inputStream.close();
	}
	
	public void addIds() throws IOException {
		String excelFilePath = "/opt/openmrs/modules/ids.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();
		
		PatientIdentifierType it = Context.getPatientService().getPatientIdentifierTypeByName("Old Identification Number");
		List<PatientIdentifierType> oldIds = new ArrayList<PatientIdentifierType>();
		oldIds.add(it);
		
		PatientIdentifierType cureId = Context.getPatientService().getPatientIdentifierTypeByUuid("81433852-3f10-11e4-adec-0800271c1b75");
		List<PatientIdentifier> identifiers = Context.getPatientService().getPatientIdentifiers(null, oldIds, null, null, null); //Do we really have to load all patients!!
		while (iterator.hasNext()) {
			try {
				Row nextRow = iterator.next();
				String patientUUID = nextRow.getCell(0).getStringCellValue();
				Cell c = nextRow.getCell(1);
				c.setCellType(Cell.CELL_TYPE_STRING);
				String idValue = c.getStringCellValue();
				if (idValue.indexOf(".") > 0)
					idValue = idValue.substring(0, idValue.indexOf("."));
				if (idValue != null && idValue.length() > 0 && !idValue.equalsIgnoreCase("NULL")) {
					for (PatientIdentifier id : identifiers) {
						if (id.getIdentifier().equalsIgnoreCase(patientUUID)) {
							PatientIdentifier identifier = new PatientIdentifier(idValue, cureId, null);
							identifier.setPatient(id.getPatient());
							Context.getPatientService().savePatientIdentifier(identifier);
						}
					}
				}
			}
			catch (Exception e) {}
			
		}
		inputStream.close();
	}
	
	public Date parseDate(String givenDate) {
		DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
		Date date = null;
		try {
			date = df.parse(givenDate);
		}
		catch (ParseException e) {
			return null;
		}
		return date;
	}
	
	public double parseToDouble(String obsValue) {
		double value = -1;
		try {
			value = Double.parseDouble(obsValue);
		}
		catch (Exception e) {
			value = -1;
		}
		return value;
	}
	
	public void saveAllPatients() {
		PatientService ps = Context.getPatientService();
		PatientIdentifierType oldIdentifierType = ps.getPatientIdentifierTypeByUuid("8d79403a-c2cc-11de-8d13-0010c6dffd0f");
		PatientIdentifierType cchuIdentifierType = ps.getPatientIdentifierTypeByUuid("81433852-3f10-11e4-adec-0800271c1b75");
		
		List<Patient> allPatients = ps.getAllPatients();
		for (Patient patient : allPatients) {
			
			Visit v = new Visit();
			v.getCreator().getDisplayString();
			
			try {
				PatientIdentifier oldId = patient.getPatientIdentifier(oldIdentifierType);
				PatientIdentifier id = patient.getPatientIdentifier(cchuIdentifierType);
				
				if (id != null && !isDuplicate(id)) {
					id.setPreferred(true);
					oldId.setPreferred(false);
				} else {
					id.setPreferred(false);
					oldId.setPreferred(true);
				}
				ps.savePatientIdentifier(id);
				ps.savePatientIdentifier(oldId);
			}
			catch (Exception e) {}
			
			ps.savePatient(patient);
		}
	}
	
	public boolean isDuplicate(PatientIdentifier id) {
		for (String duplicate : duplicateIds.split(",")) {
			if (duplicate.equalsIgnoreCase(id.getIdentifier())) {
				return true;
			}
		}
		return false;
	}
	
	public void saveAllNullDoBs() {
		
		for (Patient patient : Context.getPatientService().getAllPatients()) {
			if (patient != null && patient.getBirthdate() == null) {
				patient.setBirthdate(parseDate("1960-01-01"));
				patient.setBirthdateEstimated(true);
				Context.getPatientService().savePatient(patient);
			}
		}
		
	}
	
	public void addSurgicalProcedureObsGroup() {
		List<Encounter> encounters = new ArrayList<Encounter>();
		String[] encounterIds = Context.getAdministrationService().getGlobalProperty("cchudataimport.migratedEncounters").split(",");
		for (String string : encounterIds) {
			Encounter e = Context.getEncounterService().getEncounter(Integer.parseInt(string));
			if (e.getObs().size() > 3) { //if the encounter has more than one obs, it is an encounter from the rugical form
				encounters.add(e);
			}
		}
		ConceptService cs = Context.getConceptService();
		
		for (Encounter e : encounters) { //Every migrated encounter so fare should have an obs of concept "Surgical Procedures Form"
			Concept surgicalForm = cs.getConcept("Surgical Procedures Form");
			Obs o = new Obs(e.getPatient(), surgicalForm, e.getEncounterDatetime(), e.getLocation());
			o.setEncounter(e);
			o = Context.getObsService().saveObs(o, null);
			for (Obs obs2 : e.getAllObs()) {
				if (!obs2.getConcept().getName().toString().equalsIgnoreCase("Surgical Procedures Form")) { //Don't set obs group for the "Surgical Procedures Form" obs itself 
					o.addGroupMember(obs2);
				}
			}
			Context.getEncounterService().saveEncounter(e);
		}
		
	}
	
}
