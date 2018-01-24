import numpy as np
import xlrd as xlr
import xlwt as xlw

# This program reads a spreadsheet - see example "Decision 1.xlsx" 
# which contains AHP decision matrices
# It returns relevant AHP statistics in an output file, ex. "Decision 1 AHP Output.xlsx"
# Please contact me if you need additional documentation or help

class AHP():
    
    # Initializing AHP Object
    # Filename is read as "%filename_root%file_index%filename_ext", i.e. "Decision 1.xlsx"
    # Be sure to specify filename properly and 
    def __init__(self, file_index = None, filename_root = "Decision ", filename_ext = ".xlsx", path = None, params = None):
        self.fnr = filename_root
        self.fnx = filename_ext
        self.fi = file_index
        self.path = path
        objectives, solutions, self.params = self.load()
        self.objectives = self.oparse(objectives)
        self.solutions = self.sparse(solutions)
        self.happy_vector = None
        self.eval = None
        return
    
    # Loading Spreadsheet from file and parsing relevant parameters
    def load(self):
        if path != None:
            pass
        
        self.filename = self.fnr + str(self.fi) + self.fnx
        file = xlr.open_workbook(self.filename)
        objectives = file.sheet_by_index(0)
        solutions = file.sheet_by_index(1)
        P = int(objectives.cell(0, 3).value)
        N = int(objectives.cell(0, 4).value)

        params = {"P": P, "N": N}
        return objectives, solutions, params
    
    # Parses relevant information from objective comparison sheet    
    def oparse(self, objectives):
        P = self.params["P"]
        as_list = []
        for i in range(2, P + 2):
            as_list.append([])
            for j in range(1, P+1):
                as_list[i-2].append(objectives.cell(i, j).value)
        as_array = np.array(as_list)
        return as_array
        
    # Parses relevant information from solution comparison sheet
    def sparse(self, solutions):
        P = self.params["P"]
        N = self.params["N"]
        N0 = N + 1
        array_list = []
        for k in range(0, P):
            as_list = []
            for i in range(2, N+2):
                as_list.append([])
                for j in range(1, N + 1):
                    as_list[i-2].append(solutions.cell(i, j + N0 * k).value)
            array_list.append(as_list)
        as_array = np.array(array_list)
        return as_array
     
    # Evaluates single matrix for summary statistics
    def ahp_eval(self, matrix, q):
        cirrefs = [0, 0.52, 0.90, 1.12, 1.24, 1.32, 1.41]
        
        eigx, eigv = np.linalg.eig(matrix)
        argmax = np.argmax(eigx)
        lmax = np.amax(eigx)
        ci = (lmax - q)/(q - 1)
        cir = cirrefs[q - 2]
        cr = np.real(ci/cir)
        
        max_vec = np.abs(np.real(eigv[:,argmax]))
        max_vec /= np.sum(max_vec)
        
        return cr, max_vec
        
    # Evaluates all matrices for summary statistics
    def ahp_eval_all(self):
        obj_cr, obj_vec = self.ahp_eval(self.objectives, self.params["P"])
        sol_cr, sol_vec = [],[]
        for i in range(self.params["P"]):
            sol_cr_ret, sol_vec_ret = self.ahp_eval(self.solutions[i], self.params["N"])
            sol_cr.append(sol_cr_ret)
            sol_vec.append(sol_vec_ret)
        self.eval = {"obj": obj_cr, "sol":sol_cr, "obj_vec":obj_vec, "sol_vec": sol_vec}
        
        happy_matrix = []
        for i in range(self.params["P"]):
            happy_matrix.append(self.eval["sol_vec"][i])
        happy_matrix = np.transpose(np.array(happy_matrix))
        self.happy_vector = np.real(np.dot(happy_matrix, self.eval["obj_vec"]))
        return
    
    # Prints results to console if needed
    def disp_cr(self):
        print("The CR for the preference matrix is %f" % self.eval["obj"] + "\n The preference vector is:", self.eval["obj_vec"])
        for i in range(self.params["P"]):
            print("The CR for the evaluation matrix for objective %d is %f" % (i + 1, self.eval["sol"][i]) +  "\n The preference vector is:", self.eval["sol_vec"][i])
        
        for i in range(self.params["N"]):
            print("The overall preference for design %d is %f" % (i + 1, self.happy_vector[i]))
        print("The best design is %d." % (np.argmax(self.happy_vector) + 1))
        return

    # Ugly function that writes into a spreadsheet
    # Sheet 1 will contain CIr values and preference vectors
    # Sheet 2 will contain information on the preference for each design, as well as the best preference overall
    def write_cr(self):
        book = xlw.Workbook(encoding="utf-8")
        cr = book.add_sheet("CIr Values")
        cr.write(0,0,"CIr Value for Preference Matrix")
        cr.write(1,0, self.eval["obj"].item())
        for i in range(1, self.params["P"] + 1):
            cr.write(0,i,"CIr Value for Objective %d" % i)
            cr.write(1,i, self.eval["sol"][i - 1].item())
            
        cr.write(3, 0, "Criteria")
        cr.write(3, 1, "Normalized Preference Value")
        for i in range(self.params["P"]):
            cr.write(4 + i, 0, i)
            cr.write(4 + i, 1, self.eval["obj_vec"][i].item())
        
        
        cr.write(3, 3, "Solution Number")
        for i in range(self.params["N"]):
                cr.write(4 + i, 3, i)
        
        for i in range(self.params["P"]):
            cr.write(3, 3 + i + 1, "Normalized Preference Value for Criterion %d" % (i+1))
            for j in range(self.params["N"]):
                cr.write(4 + j, 3 + i + 1, self.eval["sol_vec"][i][j].item())

        summary = book.add_sheet("Summary")
        for i in range(self.params["N"]):
            summary.write(i, 0, "The overall preference for design %d is:" % (i+1))
            summary.write(i, 1, self.happy_vector[i].item())
            
        summary.write(self.params["N"], 0, "The best design is:")
        summary.write(self.params["N"], 1, (int(np.argmax(self.happy_vector))+1))

        book.save("%s AHP Output.xls" % (self.fnr + str(self.fi)))   

if __name__ == "__main__":
    # Sample code which reads "Decision 1.xlsx"
    # Loading from spreadsheet
    dec1 = AHP(1)
    # Evaluating all matrices
    dec1.ahp_eval_all()
    # Exporting to spreadsheet
    dec1.write_cr()
        
        
        